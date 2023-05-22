import openpyxl
import time
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Default location of the excel file to write to
excel_file_location = 'Properties.xlsx'

class ZooplaScraper:
    # Editable parameters
    location = 'London'
    min_price = '80000'
    max_price = '90000'
    
    website_url_getter = 'https://www.zoopla.co.uk/api/search/resolver/?price_frequency=per_month&results_sort=newest_listings&new_homes=include&retirement_homes=true&shared_ownership=true&include_shared_accommodation=true&search_source=for-sale&section=for-sale&view_type=list&price_max={prop_max_price}&price_min={prop_min_price}&q={prop_query1}&orig_q={prop_query2}'.format(prop_max_price=max_price, prop_min_price=min_price, prop_query1=location, prop_query2=location)
    page_number = 1
    excel_sheet = 'zoopla'
    row_index = 2

    def do_scrape(self):
        print('Scraping zoopla.co.uk')

        dr = webdriver.Chrome()
        dr.maximize_window()
        dr.get(self.website_url_getter)
        try:
            WebDriverWait(dr, 120).until(
                EC.presence_of_element_located((By.TAG_NAME, 'pre'))
                )
        except:
            print('Could not load zoopla')
            return

        website_uri = dr.find_element(By.TAG_NAME, 'pre').text

        while True:
            website_url = 'https://www.zoopla.co.uk{uri}&pn={page_number}'.format(uri=website_uri, page_number=self.page_number)
            dr.get(website_url)
            try:
                WebDriverWait(dr, 120).until(
                    EC.presence_of_element_located((By.CLASS_NAME, '_1maljyt1'))
                    )
            except:
                print('Could not load zoopla')
                break

            page_soup = BeautifulSoup(dr.page_source, 'html.parser')
            page_listings = page_soup.find_all('a', class_='_1maljyt1')

            for listing in page_listings:
                listing_url = self.convert_url(listing['href'])
                listing_desc = listing.find('div', class_='_1ankud50')
                try:
                    listing_address = listing_desc.find('h3').getText()
                except:
                    listing_address = 'Could not get address'
                try:
                    listing_title = listing_desc.find('h2').getText()
                except:
                    listing_title = 'Could not get title'
                try:
                    listing_price = listing.find('p', class_='_170k6632').getText()
                except:
                    listing_price = 'Could not get price'
                
                write_excel(self.excel_sheet, self.row_index, listing_url, listing_address, listing_title, listing_price)
                self.row_index += 1
            
            try:
                nav_menu = dr.find_element(By.CLASS_NAME, '_13wnc6k0')
                next_button = nav_menu.find_element(By.CLASS_NAME, '_1ljm00us').find_element(By.TAG_NAME, 'a')
            except:
                break
            if next_button.get_attribute('aria-disabled') == 'true':
                break

            self.page_number += 1
            #dr.close()
        
        dr.close()
    
    def convert_url(self, href):
        base_url = 'https://www.zoopla.co.uk'
        return base_url + href.split('/?')[0]

class RightMoveScraper:
    # Editable parameters
    location = 'REGION%5E305'
    min_price = '350000'
    max_price = '350000'

    website_url = 'https://www.rightmove.co.uk/property-for-sale/find.html?locationIdentifier={prop_location}&radius=0.0&maxPrice={prop_max_price}&minPrice={prop_min_price}&includeSSTC=false'.format(prop_location=location, prop_max_price=max_price, prop_min_price=min_price)
    excel_sheet = 'rightmove'
    row_index = 2
    
    def do_scrape(self):
        print('Scraping rightmove.co.uk')

        dr = webdriver.Chrome()
        dr.maximize_window()
        dr.get(self.website_url)
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
            time.sleep(1)
            try:
                WebDriverWait(dr, 20).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'propertyCard-anchor'))
                )
            except:
                print('Could not load rightome')
                break

        dr.close()

    def convert_url(self, href):
        base_url = 'https://www.rightmove.co.uk/properties/'
        return base_url + href[4:]
    
class HalmanScraper:
    # Editable parameters
    location = 'cheadle'
    min_price = '350000'
    max_price = '350000'

    website_url = 'https://www.gascoignehalman.co.uk/search/?showstc=on&showsold=on&instruction_type=Sale&place={prop_location}&ajax_border_miles=1&minprice={prop_min_price}&maxprice={prop_max_price}'.format(prop_location=location, prop_min_price=min_price, prop_max_price=max_price)
    excel_sheet = 'gascoignehalman'
    row_index = 2

    def do_scrape(self):
        print('Scraping gascoignehalman.co.uk')

        dr = webdriver.Chrome()
        dr.maximize_window()
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
                time.sleep(2)

                load_more_button = dr.find_element(By.XPATH, "//a[text()='Load More Properties']")
            except:
                break
            try:
                dr.execute_script("arguments[0].click();", load_more_button)
                time.sleep(3)
            except:
                break
        
        page_soup = BeautifulSoup(dr.page_source, 'html.parser')
        page_listings = page_soup.find_all('a', class_='btn btn-red')

        for listing in page_listings:
            listing_url = self.convert_url(listing['href'])
            write_excel(self.excel_sheet, self.row_index, listing_url)
            self.row_index += 1
    
        dr.close()

    def convert_url(self, href):
        base_url = 'https://www.gascoignehalman.co.uk'
        return base_url + href

def write_excel(sheet_name, sheet_index, url, address, title, price):
    wb = openpyxl.load_workbook(excel_file_location)
    ws = wb[sheet_name]

    ws.cell(row=sheet_index, column=1).value = url
    ws.cell(row=sheet_index, column=2).value = address
    ws.cell(row=sheet_index, column=3).value = title
    ws.cell(row=sheet_index, column=4).value = price

    wb.save(excel_file_location)
    wb.close()

def create_excel():
    global excel_file_location

    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = 'zoopla'
    sheet2 = wb.create_sheet(title='rightmove')
    sheet3 = wb.create_sheet(title='gascoignehalman')

    sheet['A1'] = 'URL'
    sheet2['A1'] = 'URL'
    sheet3['A1'] = 'URL'

    sheet['B1'] = 'Address'
    sheet2['B1'] = 'Address'
    sheet3['B1'] = 'Address'

    sheet['C1'] = 'Title'
    sheet2['C1'] = 'Title'
    sheet3['C1'] = 'Title'

    sheet['D1'] = 'Price'
    sheet2['D1'] = 'Price'
    sheet3['D1'] = 'Price'

    excel_file_location = "Properties_{sys_time}.xlsx".format(sys_time=datetime.today().strftime('%Y-%m-%d_%H-%M'))
    wb.save(excel_file_location)
    wb.close()

def main():
    create_excel()

    zoopla_scraper = ZooplaScraper()
    zoopla_scraper.do_scrape()

    """rightmove_scraper = RightMoveScraper()
    rightmove_scraper.do_scrape()

    halman_scraper = HalmanScraper()
    halman_scraper.do_scrape()"""

    print('Script ended')

if __name__ == "__main__":
    main()