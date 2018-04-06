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
    try:
        # Program starts
        start = time.time()

        date = datetime.now()
        print('Program start at:',date)
        wb = Workbook()
        ws = wb.active
        ws.append(['Date','Order','Title','Inventory','Alert Message'])
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
                'link': 'https://www.amazon.com'+item.a['href'],
                #Change string to dict
                # '{"ref":"zg_bsnr_10671048011_1","asin":"B079P3DFR4"}' is what we get first
                'asin': ast.literal_eval(item.div['data-p13n-asin-metadata'])['asin'],
                'mainImageURL': item.img['src']
            }
            products.append(product)
            # Read up to 9 (10 products)
            if index == 9:
                break
            """# Test
            if index == 1:
                break"""
            #print(product)
        # Get stock number
        # Automatic from 1-10
        #for index,product in enumerate(products):
        # Manual 
        for index,product in enumerate(products):
            product = products[index]
            print('index:',index)
            # 产品详情页
            browser.get(product['link']) 
            html = browser.page_source
            soup = BeautifulSoup(html, 'lxml')
            # Add to cart 属于产品详情页
            addToCard = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,'#add-to-cart-button')))
            addToCard.click()#this will update brower's current url
            # Go to my Cart
            cartURL = 'https://www.amazon.com'+soup.find('a',id='nav-cart')['href']
            #print("cart url",cartURL)
            browser.get(cartURL)
            html = browser.page_source
            soup = BeautifulSoup(html,'lxml')
            # Remember1:
            # Like element 'span', it doesn't support click event
            
            # BUG-Fixed不知道如何点击数量的那个按钮
            # 使用Xpath
            # Clear the input field in order to be able to send 999
            quantity = browser.find_element_by_xpath('/html/body/div[1]/div[4]/div/div[4]/div/div[2]/div[4]/form/div[2]/div/div[4]/div/div[3]/div/div/span/select/option[10]').click()
            quantity = browser.find_element_by_css_selector(".a-input-text")
            quantity.clear()
            quantity.send_keys('999')
            quantity.send_keys(Keys.RETURN)

            #BUG-Fixed不知如何得到js rendered html
            #其实js rendered html 可以配合xpath

            #enter 999后的反应时间
            time.sleep(3)
            # Get stock number from alert message if stock number is less than 999
            quantityInput = browser.find_element_by_xpath('/html/body/div[1]/div[4]/div/div[4]/div/div[2]/div[4]/form/div[2]/div/div[4]/div/div[3]/div/div/input')
            # If you wanna get a screenshot
            #browser.get_screenshot_as_file(str(index)+'.png') 
            inventory = quantityInput.get_attribute('value')
            print("how many:",inventory)
            # Add inventory attr to product
            product['inventory'] = inventory
            # Alert message
            alertMessage = browser.find_element_by_xpath('/html/body/div[1]/div[4]/div/div[4]/div/div[2]/div[4]/form/div[2]/div/div[4]/div[1]/div/div/div/span')
            alertMessageText = alertMessage.text
            product['inventoryAlertMessage'] = alertMessageText 
            # 清除库存
            emptyCart = browser.find_element_by_css_selector(".sc-action-delete > span:nth-child(1) > input:nth-child(1)")
            emptyCart.click()
        # Save to excel
        save(products,ws,wb,str(date)+'.xlsx')
        # Time used
        end = time.time()
        date = datetime.now()
        print("Ends at:",date)
        elapsed = end - start
        print("Used:",elapsed,'s')
    except Exception as err:
        print(err)
        print('库存为0')
        save(products,ws,wb,str(date)+'.xlsx')
        # Time used
        end = time.time()
        date = datetime.now()
        print("Ends at:",date)
        elapsed = end - start
        print("Used:",elapsed)

def save(products,ws,wb,wbName):
    # If KeyError: 'inventory' happens
    # Run it again will do
    for index,product in enumerate(products):
        # If you want more of the product
        #productInfo = getProductDetail(product['link'])
        # Just title and inventory
        ws.append([datetime.now(),index,product['title'],product['inventory'],product['inventoryAlertMessage']])
        #ws.append([product['title'],productInfo['starRank'],product['inventory']])
    wb.save(wbName)

def main():
    products = []
    newReleaseURL = 'https://www.amazon.com/gp/new-releases/home-garden/10671048011/ref=zg_bs_tab_t_bsnr'
    getStockNumber(newReleaseURL,products)

if __name__ == '__main__':
    main()