from pathlib import Path
import logging
import requests
import re
from urllib.parse import urljoin
from bs4 import BeautifulSoup
import panda as pd

#--logging configuration--
logging.basicConfig(
    level = logging.INFO,
    format ='%(asctime)s' | '%(levelname)-s' | '%(message)s',
    datefmt = '%H:%M:%S'
)
logger = logging.getLogger(__name__)

class PdfDownloader():
    def __init__(self,download_folder = "download_pdf"):
        # Making pdfs directory
        self.download_dir = Path(download_folder)
        self.download_dir.mkdir(exist_ok=True)

        # recording visited pdfs url
        self.history_file = Path("download_history.txt")
        self.download_url = set()

        #Load if history if it exist
        if self.history_file.exists():
            with open(self.history_file, "r", encoding="utf-8") as f:
                self.download_url = set(line.strip() for line in f if line.strip())
            logger.info(f"Loaded {len(self.download_dir)} form previously downloaded links.")

        # Making excel record
        self.master_excel_data = []
        self.master_excel_path = Path("Master_Case_summary.xlsx")
        


    def ScrapUrl(self, start_url):
        """
        fetching the url
        """
        page_to_visit = [start_url]
        visited_page = ()
        new_download_count = 0
        stop_crawling = False

        while page_to_visit and not stop_crawling:
            current_url = page_to_visit.pop(0)
            if current_url in visited_page:
                continue
            visited_page.add(current_url)
            
            logger.info(f"\n{"="*50}\n FETCHING PAGES: {current_url}\n{"="*50} ")
            try:
                response = requests.get(current_url)
                response.raise_for_status()
            except Exception as e:
                logger.error({f"Failed to fetch page from the url {current_url}"})
                continue
            
            soup = BeautifulSoup(response.content,'html.parser')


            #--- THE NEW BOX DETECTION LOGIC ---
            case_no_elements = soup.find_all(string=re.compile(r'Case No\.', re.IGNORECASE))
            if not case_no_elements:
                logger.error(f"No 'Case NO.' tag Found in {current_url}")
            else:
                logger.info(f"Found {len(case_no_elements)} case(s) on the page")

            for elements in case_no_elements:
                parent = elements.parent
                for _ in range(6):
                    if parent and parent.find('a', string=re.compile(r'PDF', re.IGNORECASE)):
                            break
                    if parent:
                        parent = parent.parent
                if not parent:
                    continue

                raw_case_no = elements.strip()
                safe_case_no = re.sub(r'[\\/*?:"<>|]', "", raw_case_no)

                date_match = re.search(r'\d{1,2}\s+[a-zA-Z]+\s+\d{4}', parent.get_text)
                case_date = date_match.group(0) if date_match else "Date not found"
                
                pdf_links = [a for a in parent.find_all('a', href=True) if a.text.upper()]

                for a_tag in pdf_links:
                    pdf_label = a_tag.text.strip()
                    safe_pdf_label = re.sub(r'[\\/*?:"<>|]', "", pdf_label)
                    print("Current URL is ", current_url)
                    print("Anchor tag is ", a_tag["href"])
                    pdf_url = urljoin(current_url, a_tag['href'])

                    if pdf_url in self.download_urls: 
                        logger.info(f"Skipping pdf alredy downloaded: {safe_case_no} - {pdf_label}")
                        continue
                    
                    pdf_filename = f"{safe_case_no} - {safe_pdf_label}"
                    logger.info(f"\n----Downloading: {pdf_filename}-----")

                    try:
                        pdf_response = requests.get(pdf_url, stream=True) 
                        pdf_response.raise_for_status()     
                    
                        file_path = self.download_dir / pdf_filename

                        with open(file_path, "wb") as f: 
                            for chunk in pdf_response.iter_content(chunk_size=8192):
                                f.write(chunk)

                        logger.info(f"Download: {file_path.name}")
                        #save to history 

                        with open(self.history_file, "a", encoding="utf-8") as f:
                            f.write(pdf_url + '\n')
                        self.download_urls.add(pdf_url)

                        # Append to master data list for the Excel file
                        self.master_excel_data.append({
                            "Case Number": safe_case_no,
                                "Date": case_date,
                                "PDF Type": pdf_label,
                                "PDF Link": pdf_url,
                                "Saved File Name": pdf_filename
                        })

                        new_download_count += 1

                        #part for pausing
                        batch = 20 
                        if new_download_count % batch == 0:
                            print(f"\n{'='*50}")
                            print(f"PAUSED: Downloaded {new_download_count} new PDFs.")
                            print("1. continue downloading")
                            print("2. Quit program (Will save Excel file first)")
                            print(f"{'='*50}")
                            ans = input("Enter 1 or 2: ").strip()
                            if ans  == "2":
                                stop_crawling = True
                                break

                    except Exception as e:
                        logger.error(f"Failed handling {pdf_url}: {e}")
            #Pagination
            if not stop_crawling:
                page_links = soup.find_all("a", class_="page-link", href=True)
                for a_tag in page_links:
                    next_page_url = urljoin(current_url, a_tag['href'].split("#")[0])
                    if next_page_url not in visited_page and next_page_url in page_to_visit:
                        page_to_visit.append(next_page_url)

            
         





def main():
    print("#"*50)
    print("please choose one of the following:")
    print("")
if __name__== "__main__":
    main()