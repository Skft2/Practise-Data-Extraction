
import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as WD
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
from datetime import datetime
import mysql.connector as connection


currentDay = datetime.now()
a = str(currentDay)
b = a.split(':')[0]
c = a.split(':')[1]

stocks = []
NOS = int(input("Enter the no of shares: "))
for i in range (NOS):
    a = input ("Enter the name of stock: ")
    stocks.append(a)

ress = []
for j in stocks:
    con = connection.connect(host='127.0.0.1',user='root',passwd='Oriental1',database='stocks')
    cur = con.cursor()
    try:
        cur.execute(f"select Symbol from nse where Company_Name like '{j}%'")
        res = cur.fetchall()
        if len(res) != 0:
            ress.append(res)
    except IndexError as e:
        pass
        
            

brand = []
for i in ress:
    brand.append(i[0][0])


ecel = []

driver =webdriver.Chrome(r'C:\Program Files\chromedriver.exe')
driver.minimize_window()
for i in brand:
    url = f"https://ticker.finology.in/company/{i}"
    driver.get(url)
    time.sleep(3)

    comname = driver.find_element_by_xpath ('//span[@id = "mainContent_ltrlCompName"]')
    sector = driver.find_elements_by_xpath('//p[@id="mainContent_compinfoId"]/strong')[2]
    cmp = driver.find_element_by_xpath ('//div[@id = "mainContent_clsprice"]//span[@class = "Number"]')
    pricechg = driver.find_element_by_xpath ('//div[@id = "mainContent_pnlPriceChange"]')
    sum = driver.find_elements_by_xpath('//div[@class="row no-gutters"]//p[@class="mb-3 mb-md-0"]')
    ess  = driver.find_elements_by_xpath('//div[@class="outputPeer w-100"]//td')

    dicts = {
        "Name" : comname.text,
        "Sector" : sector.text,
        "Current Price" : cmp.text,
        "Change in Rs" : pricechg.text.split(" ")[0],
        "Change in %" : pricechg.text.split(" ")[1].split('(')[1].split(')')[0],
        "Day's High" : sum[0].text,
        "Day's Low" : sum[1].text,
        "52-Week High" : sum[2].text,
        "52-week Low" : sum[3].text,
        "M-cap (in Cr.)" : ess[2].text,
        "P/B value" : ess[3].text,
        "P/E value" : ess[4].text,
        "EPS (in Rs.)" : ess[5].text,
        "ROE %" : ess[6].text,
        "ROCE %" : ess[7].text,
        "P/S" : ess[8].text  
    }

    ecel.append(dicts)

df = pd.DataFrame(ecel)

df.to_excel(f'MY Stock data {b}-{c}.xlsx',index=False,encoding='UTF-8')
driver.quit()    

