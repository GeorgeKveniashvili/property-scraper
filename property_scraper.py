import requests
import openpyxl
import time
from bs4 import BeautifulSoup
from selenium import webdriver

excel_file_location = "Properties.xlsx"

class RightMoveScraper:
    excel_sheet = 'rightmove'
    row_index = 2
    website_page_index = 0
    max_page_index = 0
    website_url = 'https://www.rightmove.co.uk/property-for-sale/find.html?locationIdentifier=REGION%5E87490&maxPrice=450000&minPrice=350000&index={page_index}&includeSSTC=false'
    headers = {'Sec-Ch-Ua':'Google Chrome";v="113", "Chromium";v="113", "Not-A.Brand";v="24',
               'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'}
    
    def do_scrape(self):
        print('Scraping page 1')

        page_content = requests.get(self.website_url.format(page_index=self.website_page_index), headers=self.headers)
        page_soup = BeautifulSoup(page_content.content, 'html.parser')
        page_listings = page_soup.find_all('a', class_='propertyCard-anchor')
        max_page_index = page_soup.find_all('span', class_='pagination-pageInfo')
        print(page_soup.find_all('span', {"class": "pagination-pageSelect"}))

        for listing in page_listings:
            listing_url = self.convert_url(listing['id'])
            write_excel(self.excel_sheet, self.row_index, listing_url)
            self.row_index += 1
        

    def convert_url(self, href):
        base_url = 'https://www.rightmove.co.uk/properties/'
        return base_url + href[4:]

class ZooplaScraper:
    excel_sheet = 'zoopla'
    row_index = 2
    website_page_index = 0
    max_page_index = 24
    website_url = 'https://www.zoopla.co.uk/for-sale/property/cheshire/?price_max=450000&price_min=350000&q=Cheshire&results_sort=newest_listings&search_source=home'
    
    def do_scrape(self):
        dr = webdriver.Chrome()
        dr.get(self.website_url)
        time.sleep(1)
        page_soup = BeautifulSoup(dr.page_source, 'html.parser')
        page_listings = page_soup.find_all('a', class_='_1maljyt1')

        for listing in page_listings:
            listing_url = self.convert_url(listing['href'])
            write_excel(self.excel_sheet, self.row_index, listing_url)
            self.row_index += 1
    
    def convert_url(self, href):
        base_url = 'https://www.zoopla.co.uk'
        return base_url + href
    
class HalmanScraper:
    excel_sheet = 'gascoignehalman'
    row_index = 2
    website_page_index = 0
    max_page_index = 24
    website_url = 'https://www.gascoignehalman.co.uk/search/?showstc=on&showsold=on&instruction_type=Sale&place=cheadle&ajax_border_miles=1&minprice=350000&maxprice=450000'
    
    def do_scrape(self):
        dr = webdriver.Chrome()
        dr.get(self.website_url)
        time.sleep(1)
        previous_height = dr.execute_script('return document.body.scrollHeight')

        while True:
            dr.execute_script('window.scrollTo(0, document.body.scrollHeight)')
            time.sleep(3)
            new_height = dr.execute_script('return document.body.scrollHeight')
            if new_height == previous_height:
                break
            previous_height = new_height
        
        page_soup = BeautifulSoup(dr.page_source, 'html.parser')
        page_listings = page_soup.find_all('a', class_='btn btn-red')

        for listing in page_listings:
            listing_url = self.convert_url(listing['href'])
            write_excel(self.excel_sheet, self.row_index, listing_url)
            self.row_index += 1
    
    def convert_url(self, href):
        base_url = 'https://www.gascoignehalman.co.uk'
        return base_url + href

def write_excel(sheet_name, sheet_index, url):
    wb = openpyxl.load_workbook(excel_file_location)
    ws = wb[sheet_name]

    ws.cell(row=sheet_index, column=1).value = url

    wb.save(excel_file_location)
    wb.close()

def main():
    rightmove_scraper = RightMoveScraper()
    rightmove_scraper.do_scrape()

    zoopla_scraper = ZooplaScraper()
    zoopla_scraper.do_scrape()

    halman_scraper = HalmanScraper()
    halman_scraper.do_scrape()

if __name__ == "__main__":
    main()