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

def process_pdf(pdf_path):
    print(f"\n--- Reading PDF: {pdf_path.name} ---")
    pages = convert_from_path(pdf_path, dpi=150)
    
    # List to store the data before saving to Excel
    extracted_data = []

    for i, page in enumerate(pages): 
        print(f"Processing page {i + 1}...")
        text = ocr_page(page)
        para = re.split(r'\n\s*\n', text)

        #data frame
        for p in para: 
            extracted_data.append({
                "page":i+1,
                "sentencess":p
            })
        
     
    # Save the extracted data to an Excel file
    if extracted_data:
        df = pd.DataFrame(extracted_data)
        
        # Change the file extension from .pdf to .xlsx
        excel_path = pdf_path.with_suffix('.xlsx') 
        
        # Write to Excel
        df.to_excel(excel_path, index=False)
        print(f"SUCCESS: Saved extracted text to {excel_path.name}")
    else:
        print(f"WARNING: No text could be extracted from {pdf_path.name}")

def get_title_from_pdf(pdf_path):
    print(f"Peeking into {pdf_path.name} to find a document number...")
    
    try:
        pages = convert_from_path(pdf_path, dpi=150, first_page=1, last_page=1)
        
        if pages:
            text = ocr_page(pages[0])
            
            # --- THE EXPANDED MULTI-HUNTER LIST ---
            doc_types = [
                "AB"
                "Anticipatory Bail",
                "PIL",
                "CRIL. PETITION",
                "Writ Appeal",
                "Bail Application",
                # Leaving the others here just in case you need them!
                "Civil Appeal",
                "Criminal Appeal",
                "Review Petition",
                "Special Leave Petition"
            ]
            
            for doc_type in doc_types:
                
                # --- NEW FLEXIBLE CLEANUP ---
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
                filename = f"document_{hash(pdf_url)}.pdf"# hash() this turns the url into a unique string od no 
            
            file_path = download_dir / filename
            
            with open(file_path, 'wb') as f:
                for chunk in pdf_response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            # --- NEW RENAMING & COLLISION LOGIC ---
            new_title = get_title_from_pdf(file_path)
            
            if new_title:
                # Start with the base assumption (e.g., "[Writ Appeal No. 91 of 2023].pdf")
                new_file_path = file_path.with_name(f"{new_title}.pdf")
                counter = 1
                
                # The Collision Handler: As long as a file with this name ALREADY EXISTS,
                # keep adding 1 to the counter and trying again.
                while new_file_path.exists():
                    counter += 1
                    # Try a new name: "[Writ Appeal No. 91 of 2023]2.pdf"
                    new_file_path = file_path.with_name(f"{new_title}{counter}.pdf")
                
                # Once the while loop finishes, we know we have a completely unique name.
                file_path.rename(new_file_path)
                file_path = new_file_path
                print(f"SUCCESS: Renamed file to: {file_path.name}")
            # --------------------------------------

            # Process the PDF and generate the Excel file
            process_pdf(file_path)
            
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
            process_pdf(pdf_path)
            
    elif choice == '2':
        url = input("Enter the URL to scrape: ")
        scrape_and_process_pdfs(url)
        
    else:
        print("Invalid choice. Exiting.")

if __name__ == "__main__":
    main()
