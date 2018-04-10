from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.firefox.options import Options
from selenium.webdriver import Firefox
from selenium.webdriver.common.keys import Keys
import lxml.html
#import pymongo
import re
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
import string
#Time
from datetime import datetime, date, time
import time
# Use to make a sound
import os
targetProductNameMatching = 'Maevis Bed Waterproof Mattress'

#Headless Chrome
options = webdriver.ChromeOptions()
"""options.add_argument('headless')
# Load no image makes it run faster
prefs = {"profile.managed_default_content_settings.images":2}
options.add_experimental_option("prefs",prefs)"""
options.add_argument('window-size=1200x600')
browser = webdriver.Chrome(chrome_options=options)
wait = WebDriverWait(browser, 10)
#Headless Firefox
"""options = Options()
options.add_argument('-headless')
browser = Firefox(executable_path='geckodriver', firefox_options=options)
browser.set_window_size(1200, 600)
wait = WebDriverWait(browser, 10)"""

#Excel part 
#Global variable
wb = Workbook()
products = []
adProducts = []
nonAdProducts = []
myproduct = []
def search(keyword,pageNumber,productType):
    print('正在搜索')
    # Start
    os.system('say "Your program is start now!"')
    try:
        #美亚
        browser.get('https://www.amazon.com/')
        input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#twotabsearchtextbox')))
        submit = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '.nav-search-submit > input:nth-child(2)')))
        input.send_keys(keyword)
        submit.click()
        get_products_title_index(keyword,pageNumber,productType)
    except TimeoutException:
        return search(keyword,pageNumber,productType)


def next_page(keyword,pageNumber,productType):
    print('正在翻页', pageNumber)
    try:
        wait.until(EC.text_to_be_present_in_element(
            (By.CSS_SELECTOR, '#pagnNextString'), 'Next Page'))
        submit = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#pagnNextString')))
        submit.click()
        wait.until(EC.text_to_be_present_in_element(
            (By.CSS_SELECTOR, '.pagnCur'), str(pageNumber)))
        get_products_title_index(keyword,pageNumber,productType)
        """ 位置测试 （测试环境- Google Chrome）
        # 测试发现 除了广告位会变动外
        # 自然位的不会变动 一切匹配
        # 运行正常"""
    except TimeoutException:
        next_page(keyword,pageNumber,productType)

# BUG-Fixedstring indices must be integers
# 因为转换rank前product为[] 什么都没有 
def get_products_title_index(keyword,pageNumber,productType):
    try:
        # 使用默认的排列
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#s-results-list-atf')))
        html = browser.page_source
        soup = BeautifulSoup(html, 'lxml')
        #result_0是第一个
        content = soup.find_all(attrs={"id": re.compile(r'result_\d+')})
        print("how many result were found:",len(content))
        # 如果行数超过13行 发出警报
        if len(content) > 45:
            print('More than 15 rows.')
            os.system('say "It is more than 15 rows, you should check if it is a BUG!"')
        #如果有搜索结果
        if len(content)!=0:
            for index,item in enumerate(content):
                product = {
                    #是那种非产品的缺少s-access-title的，就默认给个title
                    'title': item.find(class_='s-access-title').get_text() if item.find(class_='s-access-title') else "Amazon recommendation", 
                    'index': index+1,#在一页里的顺位序号，每一页都会变
                    #'rank': getRank(pageNumber,index),#就算有那种AD也是准的，不影响
                }
                
                products.append(product)
                # Sort product to ad and non-ad
                identifyAndSortMyProduct(product,productType)
                #Generate Rank attr for product
                turnProductIndexToRank(product,pageNumber)
                # When to stop
                if len(adProducts)>=1 and len(nonAdProducts)>=1:
                    break 
                #print(index+1," product is processed.")
        else:
            print("No products were found!")
    except Exception as err:
        print(err)

def identifyAndSortMyProduct(product,productType):
    try:
        #注意这个部分
        #亚马逊后台SKU的标题改变的话 这部分字典也要及时修改
        #防水床笠(fscl首字母缩写)部分
        fscl = {
                'Maevis Bed Waterproof Mattress Protector Cover Pad Fitted 18 Inches Deep Pocket Premium Washable Vinyl Free - Twin XL':'TXL',
                'Waterproof Mattress Cover Protector Pad with 18 Inches Deep Pocket for Full Bed by Maevis,Full Size':'F',
                'Maevis Bed Waterproof Mattress Protector Cover Pad Fitted 18 Inches Deep Pocket Premium Washable Vinyl Free - Queen':'Q',
                'Maevis Bed Waterproof Mattress Protector Cover Pad Fitted 18 Inches Deep Pocket Premium Washable Vinyl Free - King':'K',
                'Maevis Bed Waterproof Mattress Protector Cover Pad Fitted 18 Inches Deep Pocket Premium Washable Vinyl Free - California King':'CK'
            }
        #夹棉床笠部分
        jmcl = {
            'Maevis Mattress Pad Cover 100% 300TC Cotton with 8-21 Inch Deep Pocket White Overfilled Bed Mattress Topper (Down Alternative, Twin)':'T',
            'Maevis Mattress Pad Cover 100% 300TC Cotton with 8-21 Inch Deep Pocket White Overfilled Bed Mattress Topper (Down Alternative, Twin XL)':'TXL',
            'Maevis Mattress Pad Cover 100% 300TC Cotton with 8-21 Inch Deep Pocket White Overfilled Bed Mattress Topper (Down Alternative, Full)':'F',
            'Maevis Mattress Pad Cover 100% 300TC Cotton with 8-21 Inch Deep Pocket White Overfilled Bed Mattress Topper (Down Alternative, Queen)':'Q',
            'Maevis Mattress Pad Cover 100% 300TC Cotton with 8-21 Inch Deep Pocket White Overfilled Bed Mattress Topper (Down Alternative, King)':'K',
            'Maevis Mattress Pad Cover 100% 300TC Cotton with 8-21 Inch Deep Pocket White Overfilled Bed Mattress Topper (Down Alternative, California King)':'CK'
        }
        if productType == 'fscl':
            productType = fscl
        if productType == 'jmcl':
            productType = jmcl
        # Sort products to ad and nonAD
        # Title may contains blank spaces
        for matchKey in list(productType):
            if matchKey in product['title'].strip(): 
                if 'Sponsored' in product['title']:
                    adProducts.append(product)
                else:
                    nonAdProducts.append(product)
                myproduct.append(product)
                break
    except Exception as err:
        print('Identify and sorting err:',err)

def getThatTwo(productType):
    try:
        fscl = {
                'Maevis Bed Waterproof Mattress Protector Cover Pad Fitted 18 Inches Deep Pocket Premium Washable Vinyl Free - Twin XL':'TXL',
                'Waterproof Mattress Cover Protector Pad with 18 Inches Deep Pocket for Full Bed by Maevis,Full Size':'F',
                'Maevis Bed Waterproof Mattress Protector Cover Pad Fitted 18 Inches Deep Pocket Premium Washable Vinyl Free - Queen':'Q',
                'Maevis Bed Waterproof Mattress Protector Cover Pad Fitted 18 Inches Deep Pocket Premium Washable Vinyl Free - King':'K',
                'Maevis Bed Waterproof Mattress Protector Cover Pad Fitted 18 Inches Deep Pocket Premium Washable Vinyl Free - California King':'CK'
            }
            #夹棉床笠部分
        jmcl = {
            'Maevis Mattress Pad Cover 100% 300TC Cotton with 8-21 Inch Deep Pocket White Overfilled Bed Mattress Topper (Down Alternative, Twin)':'T',
            'Maevis Mattress Pad Cover 100% 300TC Cotton with 8-21 Inch Deep Pocket White Overfilled Bed Mattress Topper (Down Alternative, Twin XL)':'TXL',
            'Maevis Mattress Pad Cover 100% 300TC Cotton with 8-21 Inch Deep Pocket White Overfilled Bed Mattress Topper (Down Alternative, Full)':'F',
            'Maevis Mattress Pad Cover 100% 300TC Cotton with 8-21 Inch Deep Pocket White Overfilled Bed Mattress Topper (Down Alternative, Queen)':'Q',
            'Maevis Mattress Pad Cover 100% 300TC Cotton with 8-21 Inch Deep Pocket White Overfilled Bed Mattress Topper (Down Alternative, King)':'K',
            'Maevis Mattress Pad Cover 100% 300TC Cotton with 8-21 Inch Deep Pocket White Overfilled Bed Mattress Topper (Down Alternative, California King)':'CK'
        }
        if productType == 'fscl':
            productType = fscl
        if productType == 'jmcl':
            productType = jmcl

        targetAdRank = ''
        targetAdAttr = ''
        targetNonAdRank = ''
        targetNonAdAttr = ''
        unifiedRankAndAttr = ''

        if len(adProducts) != 0:
            targetAdRank = adProducts[0]['rank']
            # 需要将[Sponsored]移除才能匹配到 
            targetAdAttr = productType[adProducts[0]['title'].strip().replace('[Sponsored]','')] + '广告'
        if len(nonAdProducts) != 0:
            targetNonAdRank = str(nonAdProducts[0]['rank'])
            targetNonAdAttr = productType[nonAdProducts[0]['title'].strip()] + '自然'

        unifiedRankAndAttr = targetAdRank+'('+targetAdAttr+')'+'/'+targetNonAdRank+'('+targetNonAdAttr+')'
        #print("unifyied Rank",unifiedRankAndAttr)
        if unifiedRankAndAttr == "()/()":
            unifiedRankAndAttr = '大于8页'
        #获取并储存第一个广告和自然搜索的位置
        print("Two:",unifiedRankAndAttr)
        return unifiedRankAndAttr
    except Exception as err:
        print("Get that two err:",err)

# BUG-有时会出现一页超过15-20行的情况
# 按页来处理Rank
def turnProductIndexToRank(product,pageNumber):
    try:
        #Make the soup
        html = browser.page_source
        soup = BeautifulSoup(html, 'lxml')

        # 如果是九宫格和四宫格都有 默认的展示方式就可以了
        if soup.find('div',class_="s-grid-layout-picker"):
            # 选9宫格模式
            # 同时有9宫格和4宫格 
            # 都是3列 
            if soup.find('div',class_='s-image-layout-picker'):
                productIndex = product['index']
                if productIndex <= 3:
                    product['rank'] = str(pageNumber)+","+"1"+","+str(productIndex)
                # 3的倍数    
                elif productIndex%3 ==0:
                    product['rank'] = str(pageNumber)+","+str(int(productIndex/3))+","+"3"
                else:
                    product['rank'] = str(pageNumber)+","+str(int(productIndex//3 + 1))+","+str(productIndex%3)
            # 剩下的就是列模式了
            # 1_列模式可翻页的那种模式
            #如:https://www.amazon.com/s/ref=nb_sb_noss_2?url=search-alias%3Daps&field-keywords=tv&rh=i%3Aaps%2Ck%3Atv&ajr=0
            # 同时有四格和列
        elif soup.find('div',class_='s-list-layout-picker'):
            if soup.find('div',class_='s-image-layout-picker'):
                product['rank'] =  str(pageNumber)+','+str(productIndex)
        # 列
        # https://www.amazon.com/s?field-keywords=sleeping+bag
        elif not soup.find('div',class_='s-list-layout-picker'):
            if not soup.find('div',class_='s-image-layout-picker'):
                if not soup.find('div',class_="s-grid-layout-picker"):
                    if soup.find('span',id='a-autoid-0-announce').get_text() != 'See more':
                        os.system('say "Maybe it is a coloum mode, please check"')
                        print("Check please!!!")
                        product['rank'] =  str(pageNumber)+','+str(productIndex)
        else:
            #2_see more的那种像厕纸一样的中间分页的那种没有翻页的那种列模式
            #如：https://www.amazon.com/gp/vs/buying-guide/sleeping-bag/459108?ie=UTF8&field-keywords=sleeping%20bag&ref_=nb_sb_ss_ime_c_1_9&url=search-alias%3Daps
            #TODO: 什么时候解决下
            #soup.find('span',id='a-autoid-0-announce').get_text() == 'See more':
            #print("See more mode")
            #product['rank'] = "See more mode"
            #Log到Excel的rank那里表示遇到了这种情况
            print("Not the normal 3 modes")
            product['rank'] = "Other mode"
    except Exception as err:
        print("Convert to rank err:",err)

#Save Rank to Excel
def saveRankToExcel(products,pageNumber,keyword):
    try:
        for product in products:
            productTitle = product['title']
            productRank = product['rank']
            wb[keyword].append([productTitle,productRank])
        wb.save("sample.xlsx")
        print('Saved')
    except Exception as err:
        print('Save Rank failed:', err)   
        wb.save("sample.xlsx")  
    
# Title with Rank
def main():
    try:
        startTime = datetime.now()
        print("Start at:",startTime)

        global products
        # keywords to search
        #keywords = ['mattress protector','queen mattress pad','mattress topper','queen mattress topper','twin mattress pad','king mattress pad','mattress cover','mattress pad cover']
        keywords = ['mattress pad']
        productType = 'jmcl' 
        for keyword in keywords:
            #Reset pageNumber when keyword changed
            pageNumber = 1
            search(keyword,pageNumber,productType)
            #Display only the first N pages
            for pageNumber in range(2, 8):
                # When to stop turnning page
                if len(adProducts)>=1 and len(nonAdProducts)>=1:
                    break
                else:
                    next_page(keyword,pageNumber,productType)
            #重置products 如果关键词多的话 需要重置
            products = []
            print('ad:',len(adProducts))
            print('non-ad:',len(nonAdProducts))
        print(myproduct)
        #Done getting data
        getThatTwo(productType)
        endTime = datetime.now()
        print("Ends at:",endTime)
    except Exception as err:
        print('出错啦', err)
        endTime = datetime.now()
        print("Ends at:",endTime)
        #wb.save("sample.xlsx")
    finally:
        #browser.quit()
        pass
#BUG-有的界面没有那个九宫格显示模式，怎么强制切换。
#TODO:添加一个处理总时
#TODO:保存一个条目时保存一下
#TODO:'NoneType' object has no attribute 'get_text' 处理下 排查下
if __name__ == '__main__':
    main()


#BUG-出错啦 Message: Timeout loading page after 300000ms
#BUG-Fixed不能在这里退出浏览器 不然不能搜其他的产品连接了
#browser.quit()
#注意Openpyxl添加行时.append([])要添加一个list
#.append()是添加一个单元格

#TODO:sleeping bag这种Rank如何计算
#Not the normal 3 modes
#Save Rank failed: local variable 'product' referenced before assignment
#增加进度条 不然不知道是不是卡住了
#创建Github分支不要直接在Main分支操作

#非正常的3行模式product的提取方式也不一样

#BUG-EXCEL结果和网页不一样排序
#TODO：需要翻8页很浪费 处理下如何在找到2个最靠前的自然位后停止翻页

#