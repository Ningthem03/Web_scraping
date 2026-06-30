import os
import sys
import requests
import re
import logging
from bs4 import BeautifulSoup
from urllib.parse import urljoin, unquote
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
        
        # Track already downloaded URLs so we don't download them twice
        self.history_file = Path("download_history.txt")
        self.downloaded_urls = set()
        
        # Track data for the master Excel summary
        self.master_excel_data = []
        self.master_excel_path = Path("Master_Case_Summary.xlsx")
        
        # Load history if it exists
        if self.history_file.exists():
            with open(self.history_file, 'r', encoding='utf-8') as f:
                self.downloaded_urls = set(line.strip() for line in f if line.strip())
            logger.info(f"Loaded {len(self.downloaded_urls)} previously downloaded links.")

    def _save_master_excel(self):
        """Helper function to save or update the Master Excel file safely."""
        if not self.master_excel_data:
            return
            
        new_df = pd.DataFrame(self.master_excel_data)
        
        if self.master_excel_path.exists():
            # If it already exists, append the new data to the old data
            existing_df = pd.read_excel(self.master_excel_path)
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
            # Remove any duplicate rows just in case based on the specific PDF link
            combined_df.drop_duplicates(subset=['PDF Link'], inplace=True)
            combined_df.to_excel(self.master_excel_path, index=False)
        else:
            new_df.to_excel(self.master_excel_path, index=False)
            
        logger.info(f"💾 Saved {len(self.master_excel_data)} new entries to {self.master_excel_path.name}")
        # Clear the list so we don't save duplicates if this is called multiple times
        self.master_excel_data = []

    def scrape_url(self, start_url):
        """Fetches a webpage, downloads PDFs by case container, and crawls pagination links."""
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

            # --- THE BOX DETECTION LOGIC ---
            case_no_elements = soup.find_all(string=re.compile(r'Case No\.', re.IGNORECASE))
            
            if not case_no_elements:
                logger.warning("No 'Case No.' tags found on this page.")
            else:
                logger.info(f"Found {len(case_no_elements)} case(s) on this page.")
                
            for element in case_no_elements:
                if stop_crawling:
                    break

                # Traverse up to find the container/box holding the case details
                parent = element.parent
                for _ in range(6): 
                    if parent and parent.find('a', string=re.compile(r'PDF', re.IGNORECASE)):
                        break
                    if parent:
                        parent = parent.parent
                
                if not parent:
                    continue # Skip if we can't find a box containing PDFs

                # Extract Case Number and Date
                raw_case_no = element.strip()
                safe_case_no = re.sub(r'[\\/*?:"<>|]', "", raw_case_no)

                date_match = re.search(r'\d{1,2}\s+[A-Za-z]+\s+\d{4}', parent.get_text())
                case_date = date_match.group(0) if date_match else "Date Not Found"

                # Find all PDF links strictly inside THIS case's box
                pdf_links = [a for a in parent.find_all('a', href=True) if 'PDF' in a.text.upper()]
                
                for a_tag in pdf_links:
                    pdf_label = a_tag.text.strip()
                    safe_pdf_label = re.sub(r'[\\/*?:"<>|]', "", pdf_label)
                    pdf_url = urljoin(current_url, a_tag['href'])

                    if pdf_url in self.downloaded_urls:
                        logger.info(f"⏭️ Skipping already downloaded PDF: {safe_case_no} - {pdf_label}")
                        continue

                    pdf_filename = f"{safe_case_no} - {safe_pdf_label}.pdf"
                    logger.info(f"\n--- Downloading: {pdf_filename} ---")
                    
                    try:
                        pdf_response = requests.get(pdf_url, stream=True)
                        pdf_response.raise_for_status()

                        file_path = self.download_dir / pdf_filename

                        with open(file_path, 'wb') as f:
                            for chunk in pdf_response.iter_content(chunk_size=8192):
                                f.write(chunk)

                        logger.info(f"✅ Downloaded: {file_path.name}")

                        # Save to history file
                        with open(self.history_file, 'a', encoding='utf-8') as f:
                            f.write(pdf_url + '\n')
                        self.downloaded_urls.add(pdf_url)

                        # Append to master data list for the Excel file
                        self.master_excel_data.append({
                            "Case Number": safe_case_no,
                            "Date": case_date,
                            "PDF Type": pdf_label,
                            "PDF Link": pdf_url,
                            "Saved File Name": pdf_filename
                        })

                        new_downloads_count += 1

                        # Pause every 20 downloads to prevent getting blocked
                        # if new_downloads_count % 20 == 0: i plan to remove this 
                        #     print(f"\n{'='*50}")
                        #     print(f"⏸️ PAUSED: Downloaded {new_downloads_count} new PDFs.")
                        #     print("1. Continue downloading")
                        #     print("2. Quit program (Will save Excel file first)")
                        #     print(f"{'='*50}")
                        #     ans = input("Enter 1 or 2: ").strip()
                        #     if ans == '2':
                        #         stop_crawling = True
                        #         break # Break out of the PDF loop

                    except Exception as e:
                        logger.error(f"Failed handling {pdf_url}: {e}")

            # Pagination Logic
            if not stop_crawling:
                page_links = soup.find_all('a', class_='page-link', href=True)
                for a_tag in page_links:
                    next_page_url = urljoin(current_url, a_tag['href'])
                    if next_page_url not in visited_pages and next_page_url not in pages_to_visit:
                        pages_to_visit.append(next_page_url)

        # Before finishing (or quitting), always save any unsaved data to the Excel sheet
        self._save_master_excel()
        logger.info(f"\n✅ Finished! Total pages visited: {len(visited_pages)}")


def main():
    
    
    # ==========================================
    # ==========================================
    BASE_URL = "https://digihcmr.hcmdigital.in/site/index"
    TARGET_YEAR = "2023"
    TARGET_VOLUME = "2"
    TARGET_PART = ""  # Leave as "" (empty quotes) to loop through all parts 1-4
    
    print(f"--- Starting Hardcoded Scrape ---")
    print(f"Year: {TARGET_YEAR} | Volume: {TARGET_VOLUME} | Part: {TARGET_PART}")
    downloader = PDFDownloader(f"{TARGET_YEAR} Volumn_{TARGET_VOLUME}")
    
    if TARGET_PART:
        # Scrape a specific hardcoded part
        target_url = f"{BASE_URL}?VolumeViewSearch%5By_year%5D={TARGET_YEAR}&VolumeViewSearch%5Bvolume%5D={TARGET_VOLUME}&VolumeViewSearch%5Bpart%5D={TARGET_PART}"
        downloader.scrape_url(target_url)
    else:
        # If part is left blank, loop through parts 1 to 4 automatically
        for p in range(1, 5): 
            target_url = f"{BASE_URL}?VolumeViewSearch%5By_year%5D={TARGET_YEAR}&VolumeViewSearch%5Bvolume%5D={TARGET_VOLUME}&VolumeViewSearch%5Bpart%5D={p}"
            downloader.scrape_url(target_url)

if __name__ == "__main__":
    main()