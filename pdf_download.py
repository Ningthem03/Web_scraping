import os
import sys
import requests
import re
import logging
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from pathlib import Path
import pandas as pd  

# --- Configure Logging ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)-8s | %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger(__name__)

class PDFDownloader:
    def __init__(self, download_folder="downloaded_cases"):
        self.download_dir = Path(download_folder)
        self.download_dir.mkdir(exist_ok=True)
        
        # --- NEW: Create the three specific folders ---
        self.en_dir = self.download_dir / "english_judgments"
        self.rm_dir = self.download_dir / "manipuri_latin"
        self.mm_dir = self.download_dir / "manipuri_meitei"
        
        self.en_dir.mkdir(exist_ok=True)
        self.rm_dir.mkdir(exist_ok=True)
        self.mm_dir.mkdir(exist_ok=True)
        
        # Track already downloaded URLs
        self.history_file = Path("download_history.txt")
        self.downloaded_urls = set()
        
        # Track data for the master Excel summary
        self.master_excel_data = []
        self.master_excel_path = self.download_dir / "Master_Case_Summary.xlsx"
        
        # Load history
        if self.history_file.exists():
            with open(self.history_file, 'r', encoding='utf-8') as f:
                self.downloaded_urls = set(line.strip() for line in f if line.strip())
            logger.info(f"Loaded {len(self.downloaded_urls)} previously downloaded links.")

    def _save_master_excel(self):
        """Helper function to save or update the Master Excel file safely."""
        if not self.master_excel_data:
            return
            
        new_df = pd.DataFrame(self.master_excel_data)
        excel_path = Path(self.master_excel_path) 
        
        if excel_path.exists():
            existing_df = pd.read_excel(excel_path)
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
            combined_df.drop_duplicates(subset=['PDF Link'], inplace=True)
            combined_df.to_excel(excel_path, index=False)
        else:
            new_df.to_excel(excel_path, index=False)
            
        logger.info(f"💾 Saved {len(self.master_excel_data)} new entries to {excel_path.name}")
        self.master_excel_data = []

    def scrape_url(self, start_url):
        pages_to_visit = [start_url]
        visited_pages = set()
        new_downloads_count = 0
        stop_crawling = False

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
                logger.error(f"Failed to fetch page {current_url}: {e}")
                continue

            soup = BeautifulSoup(response.content, 'html.parser')

            case_no_elements = soup.find_all(string=re.compile(r'Case No\.', re.IGNORECASE))
            
            if not case_no_elements:
                logger.warning("No 'Case No.' tags found on this page.")
            else:
                logger.info(f"Found {len(case_no_elements)} case(s) on this page.")
                
            for element in case_no_elements:
                if stop_crawling:
                    break

                parent = element.parent
                for _ in range(6): 
                    if parent and parent.find('a', string=re.compile(r'PDF', re.IGNORECASE)):
                        break
                    if parent:
                        parent = parent.parent
                
                if not parent:
                    continue 

                raw_case_no = element.strip()
                
                # --- NEW: File Naming Rule - No spaces, no periods ---
                # 1. Remove standard illegal characters first
                safe_case_no = re.sub(r'[\\/*?:"<>|]', "", raw_case_no)
                # 2. Strip ALL spaces and ALL periods to create the base name
                base_name = re.sub(r'[\s\.]', '', safe_case_no) 

                date_match = re.search(r'\d{1,2}\s+[A-Za-z]+\s+\d{4}', parent.get_text())
                case_date = date_match.group(0) if date_match else "Date Not Found"

                pdf_links = [a for a in parent.find_all('a', href=True) if 'PDF' in a.text.upper()]
                
                for a_tag in pdf_links:
                    pdf_label = a_tag.text.strip()
                    pdf_url = urljoin(current_url, a_tag['href'])

                    # --- NEW: Logic to determine suffix, folder, language, and script ---
                    label_upper = pdf_label.upper()
                    
                    if "RM" in label_upper:
                        # Manipuri Roman/Latin Script
                        suffix = "mani_lat"
                        target_folder = self.rm_dir
                        language = "Manipuri"
                        script = "Latin"
                    elif "MM" in label_upper:
                        # Manipuri Meitei Script
                        suffix = "mani_meit"
                        target_folder = self.mm_dir
                        language = "Manipuri"
                        script = "Meitei"
                    else:
                        # Standard English Judgment
                        suffix = "en"
                        target_folder = self.en_dir
                        language = "English"
                        script = "Latin"

                    # Construct final filename: e.g., CaseNoWPC70of2023_en.pdf
                    pdf_filename = f"{base_name}_{suffix}.pdf"

                    if pdf_url in self.downloaded_urls:
                        logger.info(f"⏭️ Skipping already downloaded PDF: {pdf_filename}")
                        continue

                    logger.info(f"\n--- Downloading: {pdf_filename} ---")
                    
                    try:
                        pdf_response = requests.get(pdf_url, stream=True)
                        pdf_response.raise_for_status()

                        # Save to the specific language folder
                        file_path = target_folder / pdf_filename

                        with open(file_path, 'wb') as f:
                            for chunk in pdf_response.iter_content(chunk_size=8192):
                                f.write(chunk)

                        logger.info(f"✅ Saved to: {target_folder.name}/{pdf_filename}")

                        with open(self.history_file, 'a', encoding='utf-8') as f:
                            f.write(pdf_url + '\n')
                        self.downloaded_urls.add(pdf_url)

                        # --- NEW: Added 'language' and 'script' to Excel output ---
                        self.master_excel_data.append({
                            "Case Number": raw_case_no, # Keeping raw for readability in the sheet
                            "Date": case_date,
                            "PDF Type": pdf_label,
                            "PDF Link": pdf_url,
                            "Saved File Name": pdf_filename,
                            "language": language,
                            "script": script
                        })

                        # new_downloads_count += 1

                        # if new_downloads_count % 20 == 0:
                        #     print(f"\n{'='*50}")
                        #     print(f"⏸️ PAUSED: Downloaded {new_downloads_count} new PDFs.")
                        #     print("1. Continue downloading")
                        #     print("2. Quit program (Will save Excel file first)")
                        #     print(f"{'='*50}")
                        #     ans = input("Enter 1 or 2: ").strip()
                        #     if ans == '2':
                        #         stop_crawling = True
                        #         break 

                    except Exception as e:
                        logger.error(f"Failed handling {pdf_url}: {e}")

            if not stop_crawling:
                page_links = soup.find_all('a', class_='page-link', href=True)
                for a_tag in page_links:
                    # --- FIX: Removed the .split('#')[0] from pagination ---
                    next_page_url = urljoin(current_url, a_tag['href']) 
                    if next_page_url not in visited_pages and next_page_url not in pages_to_visit:
                        pages_to_visit.append(next_page_url)

        self._save_master_excel()
        logger.info(f"\n✅ Finished! Total pages visited: {len(visited_pages)}")


def main():
    downloader = PDFDownloader()
    
    # ==========================================
    # ⚙️ HARDCODED CONFIGURATION
    # ==========================================
    BASE_URL = "https://digihcmr.hcmdigital.in/site/index"
    TARGET_YEAR = "2024"
    TARGET_VOLUME = "2"
    TARGET_PART = ""  
    
    print(f"--- Starting Hardcoded Scrape ---")
    print(f"Year: {TARGET_YEAR} | Volume: {TARGET_VOLUME} | Part: {TARGET_PART}")
    
    if TARGET_PART:
        target_url = f"{BASE_URL}?VolumeViewSearch%5Bv_year%5D={TARGET_YEAR}&VolumeViewSearch%5Bvolume%5D={TARGET_VOLUME}&VolumeViewSearch%5Bpart%5D={TARGET_PART}"
        downloader.scrape_url(target_url)
    else:
        for p in range(1, 5): 
            target_url = f"{BASE_URL}?VolumeViewSearch%5Bv_year%5D={TARGET_YEAR}&VolumeViewSearch%5Bvolume%5D={TARGET_VOLUME}&VolumeViewSearch%5Bpart%5D={p}"
            downloader.scrape_url(target_url)

if __name__ == "__main__":
    main()