import decimal

import selenium
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import time
import calendar
from datetime import datetime
import datetime
from bs4 import BeautifulSoup
import scrapy
from scrapy.http import HtmlResponse
import xlsxwriter

options = webdriver.ChromeOptions()
# options.headless = True

driver = webdriver.Chrome('H:\Ian\Python Projects\web scraping\chromedriver', options=options)

# driver.maximize_window()

'''
Navigation Page
'''

# should take in pickup location, pickup date, return date
def nav(loc, pickup_date, return_date):
    try:
        driver.get("https://www.vroomvroomvroom.com.au")
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vvv-submit-search"]')))
        # pickup loc
        driver.find_element_by_id('vvv-pickup-location').send_keys(loc)
        if loc in ['AKL', 'CHC', 'WLG', 'ZQN']:
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vvv-search-form-widget"]/div[1]/div[2]/div/div[2]/div/div[2]/div[2]'))).click()
        else:
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vvv-search-form-widget"]/div[1]/div[2]/div/div[2]/div/div[2]/div[1]'))).click()
        # live in
        pickupcountry = driver.find_element_by_id('vvv-pickup-country').get_attribute('value')
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vvv-driver-residency"]'))).click()
        if pickupcountry in 'Australia':
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vvv-search-form-widget"]/div[1]/div[6]/div/div[1]/div[2]/div/div[2]/div[2]'))).click()
        else:
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vvv-search-form-widget"]/div[1]/div[6]/div/div[1]/div[2]/div/div[2]/div[7]'))).click()

        months = {'01': 'January', '02': 'February', '03': 'March', '04': 'April', '05': 'May', '06': 'June', '07': 'July',
                  '08': 'August', '09': 'September', '10': 'October', '11': 'November', '12': 'December'}
        # pickup date
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vvv-pickup-date"]'))).click()
        pickup_month = pickup_date[5:7]
        if months[pickup_month] in driver.find_element_by_id("vvv-search-form-widget").text:
            driver.find_element_by_css_selector('[data-date="' + pickup_date + '"]').click()
        else:
            while months[pickup_month] not in driver.find_element_by_id("vvv-search-form-widget").text:
                driver.find_element_by_css_selector("[title='Next Month']").click()
            driver.find_element_by_css_selector('[data-date="' + pickup_date + '"]').click()
        # return date
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vvv-return-date"]'))).click()
        return_month = return_date[5:7]
        if months[return_month] in driver.find_element_by_id("vvv-search-form-widget").text:
            driver.find_element_by_css_selector('[data-date="' + return_date + '"]').click()
        else:
            while months[return_month] not in driver.find_element_by_id("vvv-search-form-widget").text:
                driver.find_element_by_css_selector("[title='Next Month']").click()
            driver.find_element_by_css_selector('[data-date="' + return_date + '"]').click()

        # search and wait for website to show up
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vvv-submit-search"]')))
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vvv-submit-search"]'))).click()
        try:
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vvvRoot"]/div/div[6]/div[1]/div[2]/div[2]/div/div[1]/ul/li[1]/a')))
        except:
            driver.refresh()
            print('page expired, refreshing page')
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vvvRoot"]/div/div[6]/div[1]/div[2]/div[2]/div/div[1]/ul/li[1]/a')))
    except:
        print('page refresh not working, resubmitting request')
        nav(loc, pickup_date, return_date)

def scroll(speed):
    SCROLL_PAUSE_TIME = 1
    # Get scroll height
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, window.scrollY + " + str(speed) + ");")
        time.sleep(SCROLL_PAUSE_TIME)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height
# Iterate through second week for all months

def add_months(sourcedate, months):
    month = sourcedate.month - 1 + months
    year = sourcedate.year + month // 12
    month = month % 12 + 1
    day = min(sourcedate.day, calendar.monthrange(year, month)[1])
    return datetime.date(year, month, day)

def second_monday(year, month):
    c = calendar.Calendar(firstweekday=calendar.SUNDAY)
    monthcal = c.monthdatescalendar(year, month)
    try:
        first_monday = [day for week in monthcal for day in week if
                day.weekday() == calendar.MONDAY and day.month == month][0]
        second_monday = first_monday + datetime.timedelta(days=7)
        return second_monday
    except IndexError:
        print('No date found')

def month_iterator(date_list):
    # get today's date
    today = datetime.date.today()
    coming_monday = today + datetime.timedelta(days=-today.weekday(), weeks=1)
    d = datetime.timedelta(4)
    # first month
    if coming_monday.month == today.month:
        start_date = coming_monday
        return_date = start_date + d
        date_list.append([str(start_date.strftime("%Y-%m-%d")), str(return_date.strftime("%Y-%m-%d"))])
        print(start_date, return_date)
    else:
        start_date = second_monday(coming_monday.year, coming_monday.month)
        return_date = start_date + d
        date_list.append([str(start_date.strftime("%Y-%m-%d")), str(return_date.strftime("%Y-%m-%d"))])
        print(start_date, return_date)
    # following months
    i = 1
    while i < 10:
        next_month = add_months(start_date, 1)
        start_date = second_monday(next_month.year, next_month.month)
        return_date = start_date + d
        date_list.append([str(start_date.strftime("%Y-%m-%d")), str(return_date.strftime("%Y-%m-%d"))])
        i += 1
        print(start_date, return_date)

brands = ['Avis', 'Budget', 'Hertz', 'Thrifty', 'Europcar']
def price_get(brand, page_source):
    for vehicle in page_source.find_all('ul', {'class': 'vehicle-list-table list-unstyled'}):
        for lists in vehicle:
            if brand in str(lists.contents):
                print(str(brand) + ': ' + (str(lists.find('span', {'class': 'total-price-number'}))).rsplit('$', 1)[1][:-7])
                price = float(((str(lists.find('span', {'class': 'total-price-number'}))).rsplit('$', 1)[1][:-7]).replace(',', ''))
                return price

def scrape():
    locations = {
                 'SYD50': 'SYD',
                 'MEL50': 'MEL',
                 'PER50': 'PER',
                 'BRI50': 'BNE',
                 'AUC50': 'AKL',
                 'ADE50': 'ADL',
                 'COO50': 'OOL',
                 'CHR50': 'CHC',
                 'CAI50': 'CNS',
                 'DAR50': 'DRW',
                 'HOB50': 'HBA',
                 'LST50': 'LST',
                 'WEL50': 'WLG',
                 'ZQN50': 'ZQN'
                 }
    dates = []
    month_iterator(dates)
    workbook = xlsxwriter.Workbook(r'H:\中转站\vroom data ' + str(datetime.datetime.today().strftime('%Y-%m-%d')) + '.xlsx')
    worksheet = workbook.add_worksheet('vroom ' + str(datetime.datetime.today().strftime('%Y-%m-%d')))

    loc_row = 0
    loc_col = 1

    for airports in locations:
        worksheet.write(loc_row + 2, loc_col - 1, 'AVIS')
        worksheet.write(loc_row + 3, loc_col - 1, 'BUDGET')
        worksheet.write(loc_row + 4, loc_col - 1, 'HERTZ')
        worksheet.write(loc_row + 5, loc_col - 1, 'THRIFTY')
        worksheet.write(loc_row + 6, loc_col - 1, 'EUROPECAR')

        location_name = 'Location: ' + airports
        print(location_name + '-------------------------------------------------------------------------------------------------------------------------------')
        worksheet.write(loc_row, loc_col, location_name)
        date_row = loc_row + 1
        date_col = loc_col
        for date in dates:
            date_period = date[0] + ' to ' + date[1]
            print('')
            print(date_period)
            worksheet.write(date_row, date_col, date_period)

            nav(locations[airports], date[0], date[1])
            scroll(800)
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            price_row = date_row + 1
            price_col = date_col
            for brand in brands:

                brand_price = price_get(brand, soup)
                worksheet.write(price_row, price_col, brand_price)
                price_row += 1
            date_col += 1
        loc_row = date_row + 6
    workbook.close()


scrape()
driver.quit()




