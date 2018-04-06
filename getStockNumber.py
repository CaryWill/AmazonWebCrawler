#Selenium
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver import Firefox
from selenium.webdriver.common.keys import Keys
#Beautiful Soup
from bs4 import BeautifulSoup
import lxml.html
#import pymongo
import re
#Openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
#Python
import string
import ast
from datetime import datetime, date, time
import time
#From amazonCrawlers
from amazonCrawlers import getProductDetail
from amazonCrawlers import getMainImageLinks
from amazonCrawlers import getAll_Size_PriceForEachSKU
from amazonCrawlers import combine_all_size_price
from amazonCrawlers import getStarRank
from amazonCrawlers import getReviewCount
from amazonCrawlers import getAnsweredQuestionCount
#-----Done importing-------#

#Headless Chrome
options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('window-size=1200x600')
browser = webdriver.Chrome(chrome_options=options)
#Headless Firefox
"""options = Options()
#options.add_argument('-headless')
browser = Firefox(executable_path='geckodriver', firefox_options=options)
browser.set_window_size(1400, 900)"""
wait = WebDriverWait(browser, 10)

def getStockNumber(newReleaseURL,products):
    #Go the the product detail page
    try:
        browser.get(newReleaseURL)
        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,'#zg_centerListWrapper')))
        html = browser.page_source
        soup = BeautifulSoup(html, 'lxml')
        content = soup.find_all('div',class_='zg_itemWrapper')
        print('how many new release were found:',len(content))
        # Done getting new release products
        for index,item in enumerate(content):
            product = {
                'title': item.img['alt'],
                'SKUURL': 'https://www.amazon.com'+item.a['href'],
                #Change string to dict
                # '{"ref":"zg_bsnr_10671048011_1","asin":"B079P3DFR4"}' is what we get first
                'asin': ast.literal_eval(item.div['data-p13n-asin-metadata'])['asin'],
                'mainImageURL': item.img['src']
            }
            products.append(product)
            # Read up to 10 products
            if index == 10:
                break
            #print(product)
        # TODO: get detail info
        # Get stock number
        for index,product in enumerate(products):
            print('index:',index)
            for i in range(1,20):
                # 产品页
                browser.get(product['SKUURL']) 
                html = browser.page_source
                soup = BeautifulSoup(html, 'lxml')
                #print('current url1:',browser.current_url)
                #加10个产品 最后就可以不用点击输入数量了
                # Add to cart 属于产品页
                addToCard = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,'#add-to-cart-button')))
                addToCard.click()#this will update brower's current url
                # Use html = brower.page_source to update the page source
                #print('current url2:',browser.current_url)
                print("add:",i)
                if i == 10:
                    break
            # Go to my Cart
            html = browser.page_source
            soup = BeautifulSoup(html, 'lxml')
            cartURL = 'https://www.amazon.com'+soup.find('a',id='nav-cart')['href']
            #print("cart url",cartURL)
            browser.get(cartURL)
            html = browser.page_source
            soup = BeautifulSoup(html,'lxml')
            # Remember1:
            # Like element 'span', it doesn't support click event
            #----------------------------------|
            # Make quantity input field visible|
            # BUG-不知道如何点击数量的那个按钮|-----|
            #----------------------------|
            # Clear the input field in order to be able to send 999
            #browser.execute_script("arguments[0].value = arguments[1]",quantity,"0")
            #quantity = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,".a-input-text")))
            #quantity.click()
            quantity = browser.find_element_by_css_selector(".a-input-text")
            quantity.clear()
            quantity.send_keys('999')
            quantity.send_keys(Keys.RETURN)
            #BUG-不知如何得到js rendered html
            #browser.refresh()
            #html updated
            #html = browser.page_source
            #text = browser.execute_script("return document.documentElement.innerText")
            #print(text)
            #html = browser.execute_script("return document.documentElement.innerHTML")
            #soup = BeautifulSoup(html,'lxml')
            # Get stock number from alert message if stock number is less than 999
            #productDiv = soup.find('span',id='sc-subtotal-label-activecart')
            #inStock = productDiv.get_text()
            #print("how many:",inStock)
            time.sleep(3)
            browser.get_screenshot_as_file(str(index)+'.png') 
            if index == 10:
                break
    except Exception as err:
        print(err)

def main():
    products = []
    newReleaseURL = 'https://www.amazon.com/gp/new-releases/home-garden/10671048011/ref=zg_bs_tab_t_bsnr'
    getStockNumber(newReleaseURL,products)

if __name__ == '__main__':
    main()