import os
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from pathlib import Path
from pdf2image import convert_from_path
import pytesseract
from PIL import ImageOps
import pandas as pd  # Imported pandas for Excel export
import re

def ocr_page(page): 
    img = page.convert("L")
    img = ImageOps.autocontrast(img)
    text = pytesseract.image_to_string(img, lang="eng")
    return text

def get_title_from_pdf(pdf_path):
    print(f"Peeking into {pdf_path.name} to find a document number...")
    
    try:
        pages = convert_from_path(pdf_path, dpi=150, first_page=1, last_page=1)
        
        if pages:
            text = ocr_page(pages[0])
            
            # --- THE EXPANDED MULTI-HUNTER LIST ---
            doc_types = [
                "Anticipatory Bail",
                "PIL",
                "CRIL. PETITION",
                "Writ Appeal",
                "Bail Application",
                "Civil Appeal",
                "Criminal Appeal",
                "Review Petition",
                "Special Leave Petition"
            ]
            
            for doc_type in doc_types:
                
                # --- FLEXIBLE CLEANUP ---
                # 1. We make periods optional to survive "CRIL." vs "CRIL"
                # 2. We make spaces flexible to survive weird OCR spacing
                flex_type = doc_type.replace(".", r"[\.\s]*").replace(" ", r"\s*")
                
                # Build the dynamic Regex pattern
                pattern = fr'{flex_type}\s*No[\.\s]*\d+\s*of\s*\d{{4}}'
                
                match = re.search(pattern, text, re.IGNORECASE)
                
                if match:
                    # We found a match! Grab the raw text it matched.
                    raw_title = match.group(0).strip()
                    
                    # Wrap it in perfect brackets
                    safe_title = f"[{raw_title}]"
                    
                    # Sanitize it (removes any system-breaking characters)
                    safe_title = re.sub(r'[\\/*?:"<>|]', "", safe_title)
                    
                    return safe_title
                    
            print("No matching document numbers found on page 1.")
                
    except Exception as e:
        print(f"Could not read title from {pdf_path.name}: {e}")
        
    return None 

def process_pdf(pdf_path, base_excel_name):
    print(f"\n--- Reading PDF: {pdf_path.name} ---")
    pages = convert_from_path(pdf_path, dpi=150)
    
    # We will store the columns separately
    page_numbers = []
    paragraphs = []
    
    full_document_text = "" # We'll use this to check for the disclaimer

    for i, page in enumerate(pages): 
        print(f"Processing page {i + 1}...")
        text = ocr_page(page)
        full_document_text += text # Add to our giant string for checking
        
        para = re.split(r'\n\s*\n', text)

        for p in para: 
            # Only add if the paragraph isn't just empty space
            if p.strip():
                page_numbers.append(i + 1)
                paragraphs.append(p.strip())
                
    # --- THE DETECTIVE ---
    # Check if this is the translated version by looking for the keyword
    if "DISCLAIMER: Manipuri" in full_document_text or "translated in Vernacular" in full_document_text:
        print("-> Detected Vernacular/Translated Version")
        prefix = "Manipuri_"
    else:
        print("-> Detected Standard English Version")
        prefix = "English_"

    # Create a DataFrame for THIS specific PDF
    new_data = {
        f"{prefix}Page": page_numbers,
        f"{prefix}Text": paragraphs
    }
    new_df = pd.DataFrame(new_data)
    
    # --- THE PANDAS MERGE ---
    # The Excel file will be named exactly after the base case name
    excel_path = pdf_path.parent / f"{base_excel_name}.xlsx"
    
    if excel_path.exists():
        # If the Excel file already exists, open it up
        print(f"Opening existing Excel file: {excel_path.name} to merge data...")
        existing_df = pd.read_excel(excel_path)
        
        # pd.concat with axis=1 sticks the new columns side-by-side with the old ones.
        final_df = pd.concat([existing_df, new_df], axis=1)
        final_df.to_excel(excel_path, index=False)
        print(f"SUCCESS: Merged into {excel_path.name}")
    else:
        # If this is the first PDF we've processed for this case, just save it normally
        new_df.to_excel(excel_path, index=False)
        print(f"SUCCESS: Created new Excel file {excel_path.name}")

def scrape_and_process_pdfs(url):
    print(f"Fetching webpage: {url}")
    try:
        response = requests.get(url)
        response.raise_for_status() 
    except requests.exceptions.RequestException as e:
        print(f"Failed to fetch the webpage: {e}")
        return

    soup = BeautifulSoup(response.content, 'html.parser')

    # Find all PDF links
    pdf_links = set()
    for a_tag in soup.find_all('a', href=True):
        if 'PDF' in a_tag.text.upper(): 
            full_url = urljoin(url, a_tag['href'])
            pdf_links.add(full_url)

    pdf_links = list(pdf_links)
    if not pdf_links:
        print("No PDF links found on this page.")
        return

    print(f"Found {len(pdf_links)} PDF links. Starting download...")

    download_dir = Path("downloaded_pdfs")
    download_dir.mkdir(exist_ok=True)

    for i, pdf_url in enumerate(pdf_links):
        print(f"\nDownloading {i + 1}/{len(pdf_links)}: {pdf_url}")
        try:
            pdf_response = requests.get(pdf_url, stream=True)
            pdf_response.raise_for_status()
            
            filename = pdf_url.split('/')[-1]
            if not filename.lower().endswith('.pdf'):
                filename = f"document_{hash(pdf_url)}.pdf" 
            
            file_path = download_dir / filename
            
            with open(file_path, 'wb') as f:
                for chunk in pdf_response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            # --- RENAMING & COLLISION LOGIC ---
            new_title = get_title_from_pdf(file_path)
            
            if new_title:
                # Save the BASE name before appending collision numbers
                base_excel_name = new_title 
                
                new_file_path = file_path.with_name(f"{new_title}.pdf")
                counter = 1
                
                while new_file_path.exists():
                    counter += 1
                    new_file_path = file_path.with_name(f"{new_title}{counter}.pdf")
                
                file_path.rename(new_file_path)
                file_path = new_file_path
                print(f"SUCCESS: Renamed file to: {file_path.name}")
                
                process_pdf(file_path, base_excel_name)
                
            else:
                # Fallback if no legal title is found
                process_pdf(file_path, file_path.stem)
            
        except Exception as e:
            print(f"Failed to download or process {pdf_url}: {e}")

def main():
    print("Select an option:")
    print("1. Process a local PDF file")
    print("2. Scrape PDFs from a URL")
    
    choice = input("Enter 1 or 2: ")
    
    if choice == '1':
        pdf_file = input("Enter the file path (or '0' to exit): ")
        if pdf_file == '0':
            print("Program terminated") 
            exit()
            
        pdf_path = Path(pdf_file)
        if not pdf_path.exists() or pdf_path.suffix.lower() != ".pdf":
            print("INVALID: Pdf file does not exist or is not a .pdf")
        else:
            # Try to get the title for the local file too
            base_name = get_title_from_pdf(pdf_path)
            if not base_name:
                base_name = pdf_path.stem
            process_pdf(pdf_path, base_name)
            
    elif choice == '2':
        url = input("Enter the URL to scrape: ")
        scrape_and_process_pdfs(url)
        
    else:
        print("Invalid choice. Exiting.")

if __name__ == "__main__":
    main()