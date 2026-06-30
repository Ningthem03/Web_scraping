import os
import sys
import requests
import re
import logging
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin, unquote
from pathlib import Path
from pdf2image import convert_from_path
import pytesseract
from PIL import ImageOps
from PyPDF2 import PdfReader

# --- Configure Professional Logging ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)-8s | %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger(__name__)

class LegalDocumentScraper:
    def __init__(self, download_folder="downloaded_pdfs"):
        self.download_dir = Path(download_folder)
        self.download_dir.mkdir(exist_ok=True)
        
        # --- LINK TRACKER SETUP ---
        self.history_file = Path("download_history.txt")
        self.downloaded_urls = set()
        
        # Load existing history if the file exists
        if self.history_file.exists():
            with open(self.history_file, 'r', encoding='utf-8') as f:
                self.downloaded_urls = set(line.strip() for line in f if line.strip())
            logger.info(f"Loaded {len(self.downloaded_urls)} previously downloaded links from history.")
        
        # Expanded Master List of Legal Documents
        self.doc_types = [
            "Anticipatory Bail", "PIL", "CRIL. PETITION", "Writ Appeal", 
            "Bail Application", "Civil Appeal", "Criminal Appeal", 
            "Review Petition", "Special Leave Petition","Writ Petition" 
            "Writ Appeal", "PIL", "CRP", "Civil Appeal",
            "Criminal Appeal", "Review Petition", "Special Leave Petition",
            "Bail Application", "WP(C)", "WP(Crl.)", "SLP(C)", "SLP(Crl.)"
        ]

    def _hybrid_extract_text(self, pdf_path):
        """
        Attempts to read digital text natively for speed (0.1s).
        Falls back to OCR if the PDF is a scanned image (slower).
        """
        text = ""
        is_scanned = False
        
        # 1. Try Native Text Extraction (PyPDF2)
        try:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        except Exception as e:
            logger.debug(f"PyPDF2 reading failed: {e}")

        # 2. Check if text is suspiciously short/empty (meaning it's a scanned image)
        if len(text.strip()) < 100:
            is_scanned = True
            logger.info(f"Scanned document detected. Engaging OCR for {pdf_path.name}...")
            
            try:
                pages = convert_from_path(pdf_path, dpi=150)
                text = ""
                for page in pages:
                    img = ImageOps.autocontrast(page.convert("L"))
                    text += pytesseract.image_to_string(img, lang="eng") + "\n"
            except Exception as e:
                logger.error(f"OCR failed for {pdf_path.name}: {e}")
                
        return text, is_scanned


    def get_title_from_text(self, text):
        """Hunts for standard case formats and perfectly reconstructs a sanitized filename."""
        for doc_type in self.doc_types:
            # Flexible Cleanup (handles "CRIL." vs "CRIL " and bad spacing)
            flex_type = doc_type.replace(".", r"[\.\s]*").replace(" ", r"\s*")
            
            # UPGRADED PATTERN: Makes "No." and "Of" completely optional to catch "27/2023"
            pattern = fr'{flex_type}(?:\(?Civil\?)?(?:\s*NO\.?)?\s*(\d+)\s*(?:OF|/|-)?\s*(\d{{4}})'
            
            match = re.search(pattern, text, re.IGNORECASE)
            
            if match:
                # Perfectly reconstruct the title to guarantee identical file names
                case_num = match.group(1)
                case_year = match.group(2)
                return f"[{doc_type.upper()} NO. {case_num} OF {case_year}]"
                
        return None 


    def _sentence_chunker(self, text, max_words=30):
        """Splits unnumbered text into clean, sentence-aware chunks for easy reading."""
        text = re.sub(r'\s+', ' ', text).strip()
        if not text:
            return []

        sentences = re.split(r'(?<=[.!?])\s+', text)
        chunks, current, count = [], [], 0

        for s in sentences:
            words = s.split()
            if len(words) > max_words:
                if current:
                    chunks.append(" ".join(current))
                    current, count = [], 0
                for i in range(0, len(words), max_words):
                    chunks.append(" ".join(words[i:i+max_words]))
                continue

            if count + len(words) > max_words:
                chunks.append(" ".join(current))
                current, count = [s], len(words)
            else:
                current.append(s)
                count += len(words)

        if current:
            chunks.append(" ".join(current))

        return chunks

    def extract_content(self, text):
        """
        The State Machine: Finds the trigger phrase to ignore the preamble, 
        then attempts to extract sequential numbered points. 
        If no numbers are found, it extracts paragraph-by-paragraph.
        """
        # --- PREAMBLE SLICER ---
        # Look for English or Manipuri trigger phrase (case-insensitive, flexible spacing)
        trigger_pattern = r'(Judgment\s*/\s*Order\s*of\s*The\s*High\s*Court|High\s*Court\s*ki\s*Warep\s*/\s*Yathang)'
        trigger_match = re.search(trigger_pattern, text, re.IGNORECASE)
        
        if trigger_match:
            # Slice the text: keep only everything AFTER the trigger phrase
            text = text[trigger_match.end():]
            logger.info("Found trigger phrase. Slicing off preamble...")
        else:
            logger.info("No trigger phrase found. Scanning entire document...")

        # PRE-PROCESSOR: Force [1] or (1) onto a new line so we don't miss them if PyPDF2 removed line breaks
        text = re.sub(r'(?<!\n)\s*(\[\s*\d+\s*\]|\(\s*\d+\s*\))', r'\n\1', text)
        
        # Split pattern expanded to catch [1], (1), or 1.
        pattern = r'(?:^|\n)\s*(\[\s*\d+\s*\]|\(\s*\d+\s*\)|\d+\.)'
        parts = re.split(pattern, text)

        # FIXED ARRAY: Lock 20 empty slots to guarantee row alignment in Excel
        paragraphs = [""] * 20 
        expected_num = 1
        is_extracting = False
        current_chunk = ""
        current_num = 0
        found_any_numbers = False

        for part in parts:
            if re.match(r'^(\[\s*\d+\s*\]|\(\s*\d+\s*\)|\d+\.)$', part.strip()):
                point_num = int(re.sub(r'[^\d]', '', part))
                
                if not is_extracting:
                    # START CONDITION: Be flexible. If we missed [1], accept [2] or [3] as a valid start.
                    if point_num in [1, 2, 3]:
                        is_extracting = True
                        found_any_numbers = True
                        
                        # If we start at 2 or 3, the text before it was likely the missing Point 1! Save it to row 1 (index 0).
                        if point_num > 1 and current_chunk.strip():
                            paragraphs[0] = re.sub(r'\s+', ' ', current_chunk).strip()
                            
                        current_num = point_num
                        current_chunk = part.strip() + " "
                        expected_num = point_num + 1
                    else:
                        # Ignore random large numbers before we start
                        current_chunk += part
                        
                else:
                    # STOP CONDITION & FLEXIBILITY: Accept exact next number, or skip up to 2 numbers ahead.
                    if expected_num <= point_num <= expected_num + 2:
                        found_any_numbers = True
                        
                        if point_num > 20:
                            if 0 < current_num <= 20:
                                paragraphs[current_num - 1] = re.sub(r'\s+', ' ', current_chunk).strip()
                            current_chunk = ""
                            is_extracting = False
                            break
                            
                        # Save the fully built chunk explicitly to its matching row index (Point 4 -> Row 4)
                        if 0 < current_num <= 20:
                            paragraphs[current_num - 1] = re.sub(r'\s+', ' ', current_chunk).strip()
                            
                        current_num = point_num
                        current_chunk = part.strip() + " "
                        expected_num = point_num + 1
                    else:
                        # Random number out of sequence (like a [2023] citation), treat as normal text
                        current_chunk += part
                        
            else:
                # Normal text: add it to the current row (this allows accumulating preamble too)
                current_chunk += part

        # Save the very last active chunk
        if is_extracting and current_chunk.strip() and 0 < current_num <= 20:
            paragraphs[current_num - 1] = re.sub(r'\s+', ' ', current_chunk).strip()

        # --- PARAGRAPH FALLBACK ---
        # If no sequential markers found at all, fallback to sentence-aware chunking
        if not found_any_numbers:
            logger.info("No numbered points found. Falling back to sentence chunker...")
            return self._sentence_chunker(text)
            
        # Trim the empty rows at the very end of the array (stops it from printing 20 rows if only 5 exist)
        max_idx = -1
        for i in range(19, -1, -1):
            if paragraphs[i]:
                max_idx = i
                break
                
        if max_idx == -1:
            return []
            
        return paragraphs[:max_idx + 1]

    def process_file(self, pdf_path, base_excel_name=None):
        """Core logic to process a single PDF and merge it into Excel."""
        logger.info(f"Processing: {pdf_path.name}")
        
        # 1. Extract Text
        full_text, was_scanned = self._hybrid_extract_text(pdf_path)
        
        if not full_text.strip():
            logger.warning(f"Skipping {pdf_path.name}: No text extracted.")
            return

        # 2. Get Title if not provided
        if not base_excel_name:
            base_excel_name = self.get_title_from_text(full_text)
            if not base_excel_name:
                logger.warning(f"No case number found in {pdf_path.name}. Using filename.")
                base_excel_name = pdf_path.stem

        # 3. Extract Content (Points or Paragraphs)
        paragraphs = self.extract_content(full_text)

        # 4. Determine Column Name exactly based on the smart suffix we gave the file
        name_lower = pdf_path.stem.lower()
        if "mani(latin)" in name_lower:
            col_name = "Mani(latin)_Points"
        elif "mani(meitei)" in name_lower:
            col_name = "Mani(Meitei)_Points"
        elif "english" in name_lower:
            col_name = "English_Points"
        else:
            # Fallback for manually added files without explicit suffix names
            text_lower = full_text.lower()
            if "manipuri lolda" in text_lower or "translated in vernacular" in text_lower or "disclaimer: manipuri" in text_lower:
                col_name = "Mani(latin)_Points"
            else:
                col_name = "English_Points"

        # 5. Pandas Merge
        excel_path = pdf_path.parent / f"{base_excel_name}.xlsx"
        
        if excel_path.exists():
            existing_df = pd.read_excel(excel_path)
            
            # Smart logic: If we guessed Mani(latin) via text but it already exists, upgrade to Mani(Meitei)
            if col_name == "Mani(latin)_Points" and col_name in existing_df.columns:
                col_name = "Mani(Meitei)_Points"
            
            # Standard Fallback to prevent overwrites
            base_col_name = col_name
            counter = 2
            while col_name in existing_df.columns:
                col_name = f"{base_col_name}_{counter}"
                counter += 1
                
            new_df = pd.DataFrame({col_name: paragraphs})
            final_df = pd.concat([existing_df, new_df], axis=1)
            final_df.to_excel(excel_path, index=False)
            logger.info(f"MERGED -> {excel_path.name} | Added column: {col_name} ({len(paragraphs)} rows)")
        else:
            new_df = pd.DataFrame({col_name: paragraphs})
            new_df.to_excel(excel_path, index=False)
            logger.info(f"CREATED -> {excel_path.name} ({len(paragraphs)} rows)")

    def process_directory(self, folder_path):
        """Processes all PDF files in a given folder."""
        folder = Path(folder_path)
        if not folder.exists() or not folder.is_dir():
            logger.error(f"Invalid directory: {folder_path}")
            return
            
        pdf_files = list(folder.glob("*.pdf")) + list(folder.glob("*.PDF"))
        
        if not pdf_files:
            logger.warning(f"No PDFs found in the folder: {folder_path}")
            return
            
        logger.info(f"Found {len(pdf_files)} PDFs in folder. Starting bulk process...")
        for i, pdf_file in enumerate(pdf_files, start=1):
            logger.info(f"\n--- Processing File {i}/{len(pdf_files)} ---")
            self.process_file(pdf_file)
        
        logger.info(f"Finished processing {len(pdf_files)} files in folder!")

    def scrape_url(self, start_url):
        """Fetches a webpage, downloads PDFs, and automatically crawls pagination links."""
        pages_to_visit = [start_url]
        visited_pages = set()
        
        # We will store files here during Phase 1, so we can extract them all at once in Phase 2
        files_to_process = []
        new_downloads_count = 0
        stop_crawling = False
        
        # --- TITLE INHERITANCE TRACKER ---
        # Remembers the last successful case number to rescue unreadable PDFs
        last_base_title = None

        # ==========================================
        # PHASE 1: CRAWLING AND DOWNLOADING
        # ==========================================
        while pages_to_visit and not stop_crawling:
            current_url = pages_to_visit.pop(0)
            
            if current_url in visited_pages:
                continue
                
            visited_pages.add(current_url)
            logger.info(f"\n{'='*50}\n🌐 FETCHING PAGE: {current_url}\n{'='*50}")
            
            try:
                response = requests.get(current_url)
                response.raise_for_status() 
            except requests.exceptions.RequestException as e:
                logger.error(f"Failed to fetch webpage {current_url}: {e}")
                continue

            soup = BeautifulSoup(response.content, 'html.parser')

            # Collect unique PDF links while preserving top-to-bottom sequence
            raw_pdf_links = [urljoin(current_url, a['href']) for a in soup.find_all('a', href=True) if 'PDF' in a.text.upper()]
            pdf_links = []
            for link in raw_pdf_links:
                if link not in pdf_links:
                    pdf_links.append(link)

            if not pdf_links:
                logger.warning("No PDF links found on this page.")
            else:
                logger.info(f"Found {len(pdf_links)} PDF(s) on this page.")

                for i, pdf_url in enumerate(pdf_links, start=1):
                    # Check if we already downloaded this exact link in the past
                    if pdf_url in self.downloaded_urls:
                        logger.info(f"⏭️ Skipping already downloaded PDF: {pdf_url}")
                        continue
                        
                    logger.info(f"\n--- Downloading Document {i}/{len(pdf_links)} ---")
                    try:
                        pdf_response = requests.get(pdf_url, stream=True)
                        pdf_response.raise_for_status()
                        
                        # --- BULLETPROOF FILENAME SANITIZER ---
                        raw_filename = unquote(pdf_url.split('/')[-1])
                        filename = raw_filename.split('?')[0].strip()
                        filename = re.sub(r'[\\/*?:"<>|\r\n\t]', "", filename)
                        
                        if not filename.lower().endswith('.pdf'):
                            filename = f"document_{hash(pdf_url)}.pdf" 
                        
                        # GUARANTEE the folder exists right before saving 
                        self.download_dir.mkdir(parents=True, exist_ok=True)
                        temp_path = self.download_dir / filename
                        
                        # Download File
                        with open(temp_path, 'wb') as f:
                            for chunk in pdf_response.iter_content(chunk_size=8192):
                                f.write(chunk)
                        
                        # Read just enough text to find the title for neat renaming
                        temp_text, _ = self._hybrid_extract_text(temp_path)
                        base_title = self.get_title_from_text(temp_text)
                        
                        # --- INHERITANCE LOGIC: RESCUE UNREADABLE MEITEI PDFS ---
                        force_meitei = False
                        if not base_title and last_base_title:
                            base_title = last_base_title
                            force_meitei = True
                            logger.info(f"No title extracted. Inheriting previous title as Meitei script: {base_title}")
                        
                        if base_title:
                            # Keep track of this title in case the next PDF fails
                            last_base_title = base_title
                            
                            # --- SEQUENCE-BASED FILE NAMING ---
                            english_path = temp_path.with_name(f"{base_title}_English.pdf")
                            mani_latin_path = temp_path.with_name(f"{base_title}_Mani(latin).pdf")
                            
                            text_lower = temp_text.lower()
                            is_manipuri_text = "manipuri lolda" in text_lower or "translated in vernacular" in text_lower or "disclaimer: manipuri" in text_lower
                            
                            if force_meitei:
                                # Forcibly assign Meitei label if OCR failed completely
                                file_lang = "Mani(Meitei)"
                            elif not english_path.exists() and not is_manipuri_text:
                                # 1st: Standard English Version
                                file_lang = "English"
                            elif not mani_latin_path.exists():
                                # 2nd: Manipuri Latin Script
                                file_lang = "Mani(latin)"
                            else:
                                # 3rd: Manipuri Meitei Script (since English & Mani(latin) already exist)
                                file_lang = "Mani(Meitei)"

                            final_path = temp_path.with_name(f"{base_title}_{file_lang}.pdf")
                            
                            # Fallback counter just in case there are 4+ files
                            counter = 2
                            while final_path.exists():
                                final_path = temp_path.with_name(f"{base_title}_{file_lang}_{counter}.pdf")
                                counter += 1
                            
                            temp_path.rename(final_path)
                            logger.info(f"Downloaded & Renamed: {final_path.name}")
                            files_to_process.append((final_path, base_title))
                        else:
                            # Only happens if the very first PDF on the site completely fails
                            logger.info(f"Downloaded: {temp_path.name}")
                            files_to_process.append((temp_path, temp_path.stem))
                            
                        # Immediately record this successful download into the local history file
                        with open(self.history_file, 'a', encoding='utf-8') as f:
                            f.write(pdf_url + '\n')
                        self.downloaded_urls.add(pdf_url)
                        
                        new_downloads_count += 1
                        
                        if new_downloads_count % 21 == 0:
                            print(f"\n{'='*50}")
                            print(f"⏸️ PAUSED: Downloaded {new_downloads_count} new PDFs.")
                            print("What would you like to do?")
                            print("1. Continue downloading PDFs")
                            print("2. Stop downloading and start extracting now")
                            print("3. Quit program completely")
                            print(f"{'='*50}")
                            
                            while True:
                                ans = input("Enter 1, 2, or 3: ").strip()
                                if ans == '1':
                                    logger.info("Resuming downloads...")
                                    break
                                elif ans == '2':
                                    logger.info("Stopping downloads. Moving to extraction phase...")
                                    stop_crawling = True
                                    break
                                elif ans == '3':
                                    logger.info("Quitting program as requested.")
                                    sys.exit(0)
                                else:
                                    print("Invalid choice. Please enter 1, 2, or 3.")
                            
                            if stop_crawling:
                                break
                        
                    except Exception as e:
                        logger.error(f"Failed handling {pdf_url}: {e}")

            # --- PAGINATION CRAWLER ---
            if not stop_crawling:
                logger.info("Scanning for pagination links to next pages...")
                page_links = soup.find_all('a', class_='page-link', href=True)
                new_pages_found = 0
                
                for a_tag in page_links:
                    next_page_url = urljoin(current_url, a_tag['href'])
                    next_page_url = next_page_url.split('#')[0]
                    
                    if next_page_url not in visited_pages and next_page_url not in pages_to_visit:
                        pages_to_visit.append(next_page_url)
                        new_pages_found += 1
                
                if new_pages_found > 0:
                    logger.info(f"Added {new_pages_found} new page(s) to the queue. ({len(pages_to_visit)} pages remaining to scrape)")
                else:
                    logger.info("No new pagination links found on this page.")
                    
        logger.info(f"\n✅ Finished crawling all pages! Total unique web pages visited: {len(visited_pages)}")

        # ==========================================
        # PHASE 2: EXTRACTION AND EXCEL MERGE
        # ==========================================
        if files_to_process:
            logger.info(f"\n{'='*50}\n🚀 STARTING EXTRACTION PHASE ({len(files_to_process)} NEW FILES)\n{'='*50}")
            for i, (file_path, base_title) in enumerate(files_to_process, start=1):
                logger.info(f"\n--- Extracting File {i}/{len(files_to_process)} ---")
                self.process_file(file_path, base_title)
        else:
            logger.info("\nNo new PDFs were downloaded, so extraction phase is skipped.")

def main():
    scraper = LegalDocumentScraper()
    
    # Continuous Loop: Keeps asking for commands until you type '0'
    while True:
        print("\n" + "="*50)
        print("⚖️  LEGAL PDF SCRAPER & EXTRACTOR PRO ⚖️")
        print("="*50)
        print("1. Process a single local PDF file")
        print("2. Process an entire FOLDER of local PDFs (Bulk)")
        print("3. Scrape PDFs from a URL")
        print("0. Exit Program")
        
        choice = input("\nEnter your choice (0-3): ").strip()
        
        if choice == '0':
            print("Exiting program. Have a great day!")
            break
            
        elif choice == '1':
            pdf_file = input("Enter the file path: ").strip()
            pdf_path = Path(pdf_file)
            if not pdf_path.exists() or pdf_path.suffix.lower() != ".pdf":
                logger.error("Invalid file path or not a .pdf file.")
            else:
                scraper.process_file(pdf_path)
                
        elif choice == '2':
            folder_path = input("Enter the folder path (e.g., C:/downloads/pdfs/): ").strip()
            scraper.process_directory(folder_path)
                
        elif choice == '3':
            url = input("Enter the URL to scrape: ").strip()
            if url:
                scraper.scrape_url(url)
            else:
                logger.error("URL cannot be empty!")
            
        else:
            logger.error("Invalid choice. Please enter 0, 1, 2, or 3.")

if __name__ == "__main__":
    main()