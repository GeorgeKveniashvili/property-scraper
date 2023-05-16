import openpyxl
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By

excel_file_location = 'Properties.xlsx'

class RightMoveScraper:
    location = 'REGION%5E87490'
    min_price = '80000'
    max_price = '80000'
    excel_sheet = 'rightmove'
    row_index = 2
    website_url = 'https://www.rightmove.co.uk/property-for-sale/find.html?locationIdentifier={location}&maxPrice={max_price}&minPrice={min_price}&includeSSTC=false'
    
    def do_scrape(self):
        print('Scraping rightmove.co.uk')

        dr = webdriver.Chrome()
        dr.get(self.website_url.format(location=self.location, max_price=self.max_price, min_price=self.min_price))
        time.sleep(1)

        while True:
            page_soup = BeautifulSoup(dr.page_source, 'html.parser')
            page_listings = page_soup.find_all('a', class_='propertyCard-anchor')

            for listing in page_listings:
                listing_url = self.convert_url(listing['id'])
                write_excel(self.excel_sheet, self.row_index, listing_url)
                self.row_index += 1
            
            try:
                next_button = dr.find_element(By.CLASS_NAME, 'pagination-direction--next')
            except:
                break
            if not next_button.is_enabled():
                break

            dr.execute_script("arguments[0].click();", next_button)
            time.sleep(2)

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
        print('Scraping zoopla.co.uk')

        dr = webdriver.Chrome()
        dr.get(self.website_url)
        time.sleep(1)
        page_soup = BeautifulSoup(dr.page_source, 'html.parser')
        page_listings = page_soup.find_all('a', class_='_1maljyt1')

        for listing in page_listings:
            listing_url = self.convert_url(listing['href'])
            write_excel(self.excel_sheet, self.row_index, listing_url)
            self.row_index += 1
        
        dr.close()
    
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
        print('Scraping gascoignehalman.co.uk')

        dr = webdriver.Chrome()
        #dr.maximize_window()
        dr.get(self.website_url)
        time.sleep(1)
        previous_height = dr.execute_script('return document.body.scrollHeight')

        while True:
            scroll_to_elements = dr.find_elements(By.CLASS_NAME, 'btn-red')
            dr.execute_script("arguments[0].scrollIntoView();", scroll_to_elements[-1])
            time.sleep(4)
            new_height = dr.execute_script('return document.body.scrollHeight')
            if new_height == previous_height:
                break
            previous_height = new_height

        while True:
            try:
                scroll_to_elements = dr.find_elements(By.CLASS_NAME, 'btn-red')
                dr.execute_script("arguments[0].scrollIntoView();", scroll_to_elements[-1])
                time.sleep(1)

                load_more_button = dr.find_element(By.XPATH, "//a[text()='Load More Properties']")
            except:
                break
            try:
                dr.execute_script("arguments[0].click();", load_more_button)
                time.sleep(4)
            except:
                break
        
        dr.close()
        
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

    """zoopla_scraper = ZooplaScraper()
    zoopla_scraper.do_scrape()

    halman_scraper = HalmanScraper()
    halman_scraper.do_scrape()"""

if __name__ == "__main__":
    main()