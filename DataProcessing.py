from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options
from urllib.request import urlopen
import requests 
import openpyxl
import numpy as np
import time
import sys
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

def Start_yy_mm(year):
    matrix = [[str(year),'01']]
    x = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    yy    = int(x[:4])
    mm    = int(x[5:7])
    year  = int(year)
    month = 1
    if year == yy:
        while month != mm:
            month += 1
            if month < 10:
                matrix.append([str(year),'0' + str(month)])
            else:
                matrix.append([str(year),str(month)])
    else:
        matrix = [[str(year),'01'],[str(year),'02'],[str(year),'03'],[str(year),'04'],
                  [str(year),'05'],[str(year),'06'],[str(year),'07'],[str(year),'08'],
                  [str(year),'09'],[str(year),'10'],[str(year),'11'],[str(year),'12']
        ]
    return matrix

def Least_square_method(data):
    if len(data) > 6:
        data = data[-7:-1]
    x = np.arange(0,len(data))
    y = np.array(data)
    N = len(y)
    B = (sum(x[i] * y[i] for i in x) - 1./N*sum(x)*sum(y)) / (sum(x[i]**2 for i in x) - 1./N*sum(x)**2)
    A = 1.*sum(y)/N - B * 1.*sum(x)/N
    return int(A + B * N)
