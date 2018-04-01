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
targetProductNameMatching = 'Maevis Bed Waterproof Mattress'
#browser = webdriver.Firefox()
#Headless firefox config
options = Options()
options.add_argument('-headless')
browser = Firefox(executable_path='geckodriver', firefox_options=options)
wait = WebDriverWait(browser, 10)
browser.set_window_size(1400, 900)

#Excel part
wb = Workbook()
products = []

def search(keyword,pageNumber,sheetNumber,worksheet,myProductIDForMatcting):
    print('正在搜索')
    try:
        #美亚
        browser.get('https://www.amazon.com/')
        input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#twotabsearchtextbox')))
        submit = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '.nav-search-submit > input:nth-child(2)')))
        input.send_keys(keyword)
        submit.click()
        get_products(keyword,pageNumber,sheetNumber,worksheet,myProductIDForMatcting)
    except TimeoutException:
        return search(keyword,pageNumber,sheetNumber,worksheet,myProductIDForMatcting)


def next_page(keyword,pageNumber,sheetNumber,worksheet,myProductIDForMatcting):
    print('正在翻页', pageNumber)
    try:
        wait.until(EC.text_to_be_present_in_element(
            (By.CSS_SELECTOR, '#pagnNextString'), 'Next Page'))
        submit = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#pagnNextString')))
        submit.click()
        wait.until(EC.text_to_be_present_in_element(
            (By.CSS_SELECTOR, '.pagnCur'), str(pageNumber)))
        get_products(keyword,pageNumber,sheetNumber,worksheet,myProductIDForMatcting)
    except TimeoutException:
        next_page(keyword,pageNumber,sheetNumber,worksheet,myProductIDForMatcting)


def get_products(keyword,pageNumber,sheetNumber,worksheet,myProductIDForMatcting):
    try:
        #切换为3列模式，不然数据是列模式的。
        #condition1:有9宫格按钮
        #threeColoumMode = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.s-grid-layout-picker')))
        #condition2:只有四格按钮
        #threeColoumMode = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'span.s-layout-toggle-picker:nth-child(3)'))) 
        #threeColoumMode.click()
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#s-results-list-atf')))
        html = browser.page_source
        soup = BeautifulSoup(html, 'lxml')
        content = soup.find_all(attrs={"id": re.compile(r'result_\d+')})
        #获取产品链接
        #利用图片的父node来提取链接
        urls = []
        productDivs = soup.findAll('img', attrs={'class':'s-access-image cfMarker'})
        #print("links:",len(productDivs))
        for div in productDivs:
            urls.append(div.parent['href'])
        #如果有搜索结果
        if len(content)!=0:
            for (item,index,url) in zip(content,range(1,9999),urls):
                try:
                       item.find(class_='s-access-title').get_text() 
                except AttributeError:
                        print('title error')
                product = {
                    #是那种非产品的缺少s-access-title的，就默认给个title
                    #此处的get_text不会引 'NoneType' object has no attribute 'get_text'
                    
                    'title': item.find(class_='s-access-title').get_text() if item.find(class_='s-access-title') else "Amazon recommendation", 
                    'index': index,
                    #'rank': getRank(pageNumber,index),#就算有那种AD也是准的，不影响
                    #'image':item.find(class_='s-access-image cfMarker').get('src')
                }
                product['link'] = 'https://www.amazon.com'+url if ('Sponsored' in product['title']) else url
                products.append(product)
                #wb[keyword].append([product['title'],product['rank'],product['link']])
    except Exception as e:
        print(e)

def getRank(pageNumber,productIndex):
    if productIndex <= 3:
        return str(pageNumber)+","+"1"+","+str(productIndex)
    elif productIndex%3 ==0:
        #3的倍数肯定是第三列
        #而且是整数列  
        return str(pageNumber)+","+str(productIndex//3)+","+"3"
    else:
        return str(pageNumber)+","+str(productIndex//3 + 1)+","+str(productIndex%3)


#read excel file and process&format the file
def processExcel():
    wb = load_workbook(filename = "sample.xlsx")
    wb['Sheet'].append(['Good morning!'])
    wb.save(filename = "sample.xlsx")

def main():
    try:
        # keywords to search
        #keywords = ['queen mattress protector','king mattress protector','waterproof mattress pad']
        keywords = ['mattress protector']
        for (keyword,sheetx) in zip(keywords,range(1,999)):
            global products
            # one keyword per sheet
            ws = wb.create_sheet(title=keyword)
            # initial it with one row
            ws.append(["Product Name", "Star Rank","Review Count","SKU price","Main image link","Product Link"])
            #Reset pageNumber when keyword changed
            pageNumber = 1
            search(keyword,pageNumber,sheetx,ws,targetProductNameMatching)
            #Display only the first N pages
            for pageNumber in range(2, 2):
                next_page(keyword,pageNumber,sheetx,ws,targetProductNameMatching)
            #Done getting all products
        #Process all prodcuts obtained
        #Open single product link to get product detail
        #print(len(products))
        #注意Openpyxl添加行时.append([])要添加一个list
        #.append()是添加一个单元格
        #TODO:size没处理好
        for productIndex,product in enumerate(products):
            productsDetailInfoDict = getProductDetail(product['link'])
            #process sku size
            skuSizes = productsDetailInfoDict['size']
            skuSizesCombination = ""
            #[{}, {'Twin': '$24.95'}, {'Twin XL': '$24.95'}, {'Full': '$25.95'}, {'Queen': '$26.95'}, {'King': '$34.95'}, {'California King': '$34.95'}] 
            # BUG-会出现{},是哪里抓取出问题了吗
            if len(skuSizes) != 0:
                for index,skuSizeDict in enumerate(skuSizes):
                    if bool(skuSizeDict):#如果不是空字典
                        key = list(skuSizeDict)[0]
                        value = skuSizeDict[key]
                        maxIndex = len(skuSizes)-1
                    if index < maxIndex:#最后一个不用加/
                        skuSizesCombination += key+":"+value+'/'
                    else:
                        skuSizesCombination += key+":"+value
            #print(skuSizesCombination)
            #ws.append([product['title'],skuSizesCombination, productsDetailInfoDict['starRank'],productsDetailInfoDict['reviewCount'],productsDetailInfoDict['imageLink'],product['link']])
            ws.append([product['title'],skuSizesCombination, productsDetailInfoDict['starRank'],productsDetailInfoDict['imageLink'],product['link']])
            print('Saved',productIndex)
            #print(productsDetailInfoDict)
        #重置products
        products = []
        wb.save("sample.xlsx")
        #browser.quit()
    except Exception as e:
        print('出错啦', e)
        wb.save("sample.xlsx")
    finally:
        browser.quit()
        #pass
#-------------Test---------------
#'NoneType' object has no attribute 'get_text'
#出现在这个getProductDetail()函数里
def getProductDetail(productLink):
    try:
        #BUG:速度是不是会很慢如果用了很多的find_all() 
        browser.get(productLink) 
        html = browser.page_source
        soup = BeautifulSoup(html, 'lxml') 
        
            """try:
                soup.find('span',id='priceblock_ourprice').get_text() 
            except AttributeError:
                print("price our error")"""
            """try:
            soup.find('span',id='priceblock_dealprice').get_text()
            except AttributeError:
                print("deal price error")"""
            
            #print(sizes)
            #print(key,price)
            #BUG-Fixed不能在这里退出浏览器 不然不能搜其他的产品连接了
            #browser.quit()
        #Get star rank
        starTag = soup.find('i',class_="a-icon a-icon-star a-star-4-5")
        #找评论和Q&A一起的div
        #starTag = soup.find('div',id='averageCustomerReviews')#.find('span',id='acrPopover')
        #starRank = starTag.span['title']
        starRank = starTag.span.get_text()
        #此处的get_text不会引 'NoneType' object has no attribute 'get_text'
        #old way
        try:
            starRank = starTag.span.get_text()
        except AttributeError:
            print("star tag error")
        #print("star:",starRank)
        #有两个地方会显示评分而且这两个一样 所以用find就够了
        #此处的get_text不会引 'NoneType' object has no attribute 'get_text'
        reviewTag = soup.find('span',id = 'acrCustomerReviewText')
        #reviewTag = soup.find('div',id='averageCustomerReviews').find('span',id='acrCustomerReviewText')
        reviewCount = reviewTag.span.get_text()
        #old ways
        try:
            reviewCount = reviewTag.span.get_text()
        except AttributeError:
            print("reivew tag error")
        #print(reviewCount)
        #Get Q&A 
        #此处的get_text不会引 'NoneType' object has no attribute 'get_text'
        #TODO：测试
        #修改下find div->tag1->attr
        #这样限制范围是不是会提速
        answeredQuestionTag = soup.find('a', id="askATFLink")
        try:
            answeredQuestionCount = answeredQuestionTag.span.get_text().strip()
        except AttributeError:
            print("Answer tag error")    
        #print(answeredQuestionCount)
        #print(sizes)
        return {'size':sizes,'price':price,'starRank':starRank,'reviewCount':reviewCount,'imageLink':imageLinks[0]}
        #test
        #return {'size':sizes,'price':price,'starRank':starRank,'imageLink':imageLinks[0]}
    except Exception as e:
        print('getProDeatil', e)


def getMainImageLinks(soup):
    #主图
    imageLinks = []
    #BUG-不知道如何提取hidden的itemNo-去Stack Overflow上问了
    #难怪各种方式提取 都只有一个li tag被提取出来
    imageTags = soup.find_all('li',class_=re.compile(r'itemNo'))
    for image in imageTags:
        imageLinks.append(image.img['src'])
    return imageLinks
    
def getPriceForEachSKU(soup):
    #BUG-Fixed为什么(r'size_name_')而不是(r'size_name_\b+')因为是\d+不是\b+
    sizeTags = soup.find_all('li',id=re.compile(r'size_name_\d+'))
    #每个SKU的价格
    #Get all sizes
    sizes = [{}]
    for size in sizeTags:
        key = str.replace(size['title'],'Click to select ','')
        skuLink = 'https://www.amazon.com'+size['data-dp-url']
        if skuLink == 'https://www.amazon.com':
            skuLink = productLink
        #open new tab in browser to get the price
        browser.get(skuLink)
        html = browser.page_source
        soup = BeautifulSoup(html, 'lxml')
        #选中价格时那这个价格就是这一页的价格 因为skulink提取出来是""
        #如果价格是那种有折扣的话 id = "priceblock_dealprice"
        #此处的两个get_text不会引 'NoneType' object has no attribute 'get_text' double checked
        price = soup.find('span',id='priceblock_ourprice').span.get_text() if soup.find('span',id='priceblock_ourprice') else soup.find('span',id='priceblock_dealprice').span.get_text()
        #print(price) 
        sizes.append({key:price})
#-------------End--------------------  
#BUG-有的界面没有那个九宫格显示模式，怎么强制切换。
#TODO:添加一个处理总时间
#TODO:保存一个条目时保存一下
#TODO:'NoneType' object has no attribute 'get_text' 处理下 排查下
if __name__ == '__main__':
    main()


#BUG-出错啦 Message: Timeout loading page after 300000ms
