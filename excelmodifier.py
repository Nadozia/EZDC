import openpyxl
import pandas as pd
import requests
from bs4 import BeautifulSoup
import datetime
import os

class ExcelProcessor:
    def __init__(self, filepath, product_sheet_name='Sheet1'):
        self.filename = filepath
        try:
            self.wb = openpyxl.load_workbook(filename=filepath)
        except:
            self.wb = openpyxl.load_workbook(filename='./EZDC.xlsx')
        self.product_sheet_name = product_sheet_name
        self.sheet_names = self.wb.sheetnames
        self.product_list = self.getProducts()
        self.session = requests.session()
                

    def getProducts(self):
        ws = self.wb[self.product_sheet_name]
        products = []
        for r in range(ws.max_row):
            products.append(ws.cell(row=r+1, column=4).value)
        return products[1:]

    def updateWorkSheets(self):
        for product in self.product_list:
            product_asin = str(product)
            if product_asin not in self.sheet_names:
                product_ws = self.wb.create_sheet(title=product_asin)
                product_ws.append(('Date', 'ASIN', 'STARS', 'NUM_RATING', 'RANKING1','RANKING1_CAT','RANKING2', 'RANKING2_CAT'))
            else:
                product_ws = self.wb[product_asin]
            Date = datetime.datetime.today()
            try:
                ASIN, REVIEW, REVIEW_TEXT, rankings = self.getStarsReview(session=self.session, ASIN=product_asin)
                product_ws.append((Date, ASIN, REVIEW, REVIEW_TEXT, rankings[0][0], rankings[0][1], rankings[1][0], rankings[1][1]))
                print(f'Append Successful on {ASIN}.')
                print((Date, ASIN, REVIEW, REVIEW_TEXT, rankings))
            except:
                print(f'No page for {ASIN}.')
                product_ws.append((Date, product_asin, 'NAN', 'NAN', 'NAN', 'NAN', 'NAN', 'NAN'))
        
        self.wb.save(self.filename)

    def getStarsReview(self, session, ASIN):
        HEADERS = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:66.0) Gecko/20100101 Firefox/66.0", "Accept-Encoding":"gzip, deflate", "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8", "DNT":"1","Connection":"close", "Upgrade-Insecure-Requests":"1"}
        url = 'https://www.amazon.com/dp/'+ASIN
        r = session.get(url, headers=HEADERS)
        s = BeautifulSoup(r.content, 'lxml')
        result = s.find_all('table', {'class':'prodDetTable'})
        ASIN = 'CHECKED EMPTY'
        REVIEW = 'CHECKED EMPTY'
        REVIEW_TEXT = 'CHECKED EMPTY'
        rankings = []
        for t in result:
            entries = t.find_all('th',{'class':'prodDetSectionEntry'})
            for ent in entries:
                if 'ASIN' in ent.text:
                    try:
                        ASIN = ent.nextSibling.nextSibling.string.replace('\n','')

                    except:
                        pass
                if 'Reviews' in ent.text:
                    try:
                        REVIEW = float(ent.nextSibling.nextSibling.find('span',{'id':'acrPopover'})['title'].replace(',','').replace(' out of 5 stars',''))
                        REVIEW_TEXT = int(ent.nextSibling.nextSibling.find('span',{'id':'acrCustomerReviewText'}).string.replace(',','').replace('\n','').replace(
                        ' ratings',''))
                    except Exception as e:
                        print(e)
                if 'Rank' in ent.text:
                    try:
                        RANKING = ent.nextSibling.nextSibling
                        for r in RANKING.find_all('span'):
                            text = r.text
                            if '(' in text:
                                text = text[:text.index('(')]
                            texts = text.replace('#','').replace(r'\n','').replace('&amp;','&').split(' in ')
                            rank_num = int(texts[0].replace(',','').strip())
                            rank_cat = texts[1].strip()
                            rankings.append((rank_num, rank_cat))
                    except:
                        rankings = [('CHECKED EMPTY','CHECKED EMPTY'),('CHECKED EMPTY','CHECKED EMPTY')]
                 
        return(ASIN, REVIEW, REVIEW_TEXT, rankings[1:])

def Main():
    filename = './EZDC.xlsx'
    mtime = os.path.getmtime(filename)
    EP = ExcelProcessor(filepath=filename)
    if datetime.datetime.today().date()!=datetime.datetime.utcfromtimestamp(mtime).date():
        EP.updateWorkSheets() 
    else:
        print(f'{filename} has been modified today, can you sure to update?(y/n)')
        ans = str(input())
        if ans == 'y' or ans == 'Y':
            EP.updateWorkSheets() 
        else:
            print('Bye~')

if __name__ =='__main__':
    Main()