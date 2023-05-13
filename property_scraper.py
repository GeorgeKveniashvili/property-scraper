import requests
import openpyxl
from bs4 import BeautifulSoup

excel_file_location = "Properties.xlsx"


class RightMoveScraper:
    excel_sheet = "rightmove"
    row_index = 2
    website_page_index = 0
    max_page_index = 24
    website_url = 'https://www.rightmove.co.uk/property-for-sale/find.html?locationIdentifier=REGION%5E87490&maxPrice=450000&minPrice=350000&index={page_index}&includeSSTC=false'
    headers = {'Sec-Ch-Ua':'Google Chrome";v="113", "Chromium";v="113", "Not-A.Brand";v="24',
               'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'}
    
    def do_scrape(self):
        print('Scraping page 1')

        rightmove_page = requests.get(self.website_url.format(page_index=self.website_page_index), headers=self.headers)
        rightmove_soup = BeautifulSoup(rightmove_page.content, 'html.parser')
        rightmove_listings = rightmove_soup.find_all('a', class_='propertyCard-anchor')
        max_page_index = rightmove_soup.find_all('span', class_='pagination-pageInfo')
        print(rightmove_soup.find_all('span', {"class": "pagination-pageSelect"}))

        for listing in rightmove_listings:
            listing_url = self.convert_url(listing['id'])
            write_excel('rightmove', self.row_index, listing_url)
            self.row_index += 1
        

    def convert_url(self, anchor_id):
        base_url = 'https://www.rightmove.co.uk/properties/'
        return base_url + anchor_id[4:]


def write_excel(sheet_name, sheet_index, url):
    wb = openpyxl.load_workbook(excel_file_location)
    ws = wb[sheet_name]

    ws.cell(row=sheet_index, column=1).value = url

    wb.save(excel_file_location)
    wb.close()

def main():
    right_move_scraper = RightMoveScraper()
    right_move_scraper.do_scrape()

if __name__ == "__main__":
    main()





#zoopla_url = "https://www.zoopla.co.uk/for-sale/property/cheshire/?price_max=450000&price_min=350000&q=Cheshire&results_sort=newest_listings&search_source=home"

"""zoopla_page = requests.get(zoopla_url, headers=headers)
print(zoopla_page)

zoopla_soup = BeautifulSoup(zoopla_page.content, 'html.parser')

zoopla_listings = zoopla_soup.find_all('a', class_='listing-results-price text-price')"""