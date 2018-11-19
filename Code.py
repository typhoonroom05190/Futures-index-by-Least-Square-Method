from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from urllib.request import urlopen
import requests     # requests.post
import openpyxl
import numpy as np
import time
import sys
import os
import re

def Check_monthly_closing_price (year,month):

    current_year  = int(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())[:4])
    current_month = int(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())[5:7])

    if current_year == int(year) and current_month == int(month):
        base_url = 'http://www.twse.com.tw/zh/page/trading/indices/MI_5MINS_HIST.html'
        chrome_options = Options()
        chrome_options.add_argument("--headless")       # define headless
        driver = webdriver.Chrome(chrome_options=chrome_options)  
        driver.get(base_url)
        html = driver.page_source
        driver.close()
        soup = BeautifulSoup(html, features='lxml')
        path = soup.find_all('a',{'class':re.compile('html')})
        url = 'http://www.twse.com.tw' + path[0]['href']
        html = requests.get(url)
        html = urlopen(url).read().decode('utf-8')
        soup = BeautifulSoup(html, features='lxml')
        content = soup.find_all('td')[5:]
        closing_price = content[4::5]
        date = content[0::5]
    else:
        base_url = 'http://www.twse.com.tw/indicesReport/MI_5MINS_HIST?response=html&date='
        url = (base_url + year + month + '01')
        html = requests.get(url)
        html = urlopen(url).read().decode('utf-8')
        soup = BeautifulSoup(html, features='lxml')
        content = soup.find_all('td')[5:]
        closing_price = content[4::5]
        date = content[0::5]

    data = dict()
    for i in range(len(date)):
        price = float(closing_price[i].get_text().replace(',',''))
        day   = date[i].get_text()
        data[day] = price
    return date, data

def Start_yy_mm(year,month):
    matrix = []
    x = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    yy    = int(x[:4])
    mm    = int(x[5:7])
    year  = int(year)
    month = int(month)
    while year != yy or month != mm:
        if month < 10:
            matrix.append([str(year),'0' + str(month)])
        else:
            matrix.append([str(year),str(month)])
        month += 1
        if month > 12:
            year += 1
            month = 1
    if month < 10:
        matrix.append([str(year),'0' + str(month)])
    else:
        matrix.append([str(year),str(month)])
    return matrix

def Least_square_method(data):
    if len(data) == 1:
        return int(data[0])
    if len(data) > 6:
        data = data[-7:-1]
    x = np.arange(0,len(data))
    y = np.array(data)
    N = len(y)
    B = (sum(x[i] * y[i] for i in x) - 1./N*sum(x)*sum(y)) / (sum(x[i]**2 for i in x) - 1./N*sum(x)**2)
    A = 1.*sum(y)/N - B * 1.*sum(x)/N
    return int(A + B * N)

name = input('請輸入你的檔案名稱:')
try:
    wb = openpyxl.load_workbook(name + '.xlsx')
except Exception as X:
    print('找不到該檔案')
    response = input('請問你要新建一個檔案嗎?\n回答:')
    if response == 'yes':
        wb = openpyxl.Workbook()
        wb.save(name + '.xlsx')
        openpyxl.load_workbook(name + '.xlsx')
    else:
        sys.exit()

year = input('請輸入你要的年份:')
month = input('請輸入你要的月份:')
current_year  = int(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())[:4] )
current_month = int(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())[5:7])
while current_month == int(month) and current_year == int(year):
    month = input('資料量過少，請重新輸入月份:')

total = Start_yy_mm(year,month)

share_price = []
discriminate = ['賣出']
index = [0,0,0]

for t in total:
    date, data = Check_monthly_closing_price(t[0],t[1])
    sheet = wb['Sheet']
    x = sheet.max_row
    for d in date:
        share_price.append(float(data[d.get_text()]))
        sheet['A'+str(x)] = d.get_text()
        sheet['B'+str(x)] = int(data[d.get_text()])
        if len(share_price) > 6:
            sheet['C'+str(x)] = int(Least_square_method(share_price))
            sheet['D'+str(x)] = int(data[d.get_text()]) - int(Least_square_method(share_price))

            index.append(int(data[d.get_text()]) - int(Least_square_method(share_price)))
            if (index[-2] < 0 < index[-1] or 0 < index[-2] < index[-1]) or (index[-3] < 0 and index[-2] < index[-1] and -30 < index[-1] < 0):
                if discriminate[-1] == '買進':
                    pass
                else:
                    sheet['E'+str(x)] = '買進'
                    discriminate.append('買進')
            elif (index[-2] > 0 > index[-1] or 0 > index[-2] > index[-1]) or (index[-3] > 0 and index[-2] > index[-1] and 0 < index[-1] < 30):
                if discriminate[-1] == '賣出':
                    pass
                else:
                    sheet['E'+str(x)] = '賣出'
                    discriminate.append('賣出')
        x += 1
        wb.save(name + '.xlsx')
