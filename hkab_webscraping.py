#webscraping
#python3.7

import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
from pandas import DataFrame
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from datetime import date, timedelta

########################################################
#Date
def date_func(yyyy,mm,d):
    year = driver.find_element_by_xpath('//option[contains(@value, ' + yyyy + ')]')
    year.click()
    month = driver.find_element_by_xpath('//option[contains(@value, ' + mm + ')]')
    month.click()
    day = driver.find_element_by_xpath('//a[contains(@href,"javascript:isSubmit(\'' + d + '\')")]')
    day.click()

#Table 1
def tb1_func():
    # heading
    head = []
    try:
        elements = driver.find_elements_by_xpath(
            '//*[@id="testTable"]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td')
        for each in elements:
            head.append(each.text)
        elements = driver.find_elements_by_xpath(
            '//*[@id="testTable"]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr')
        list = []
        for each in elements:
            list.append(each.text)
        Row = len(list) - 1
    except:
        None
    # data_table 1
    tb1_list = []
    tb1_list.append(head)
    row = []
    for r in range(4, Row + 4):
        try:
            elements = driver.find_elements_by_xpath(
                '//*[@id="testTable"]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr['+str(r)+']/td')
            for each in elements:
                try:
                    ele = float(each.text)
                except:
                    ele = each.text
                row.append(ele)
        except:
                None
        tb1_list.append(row)
        row = []
    return tb1_list

#data_table 2
def tb2_func():
    head = []
    try:
        elements = driver.find_elements_by_xpath(
            '//*[@id="testTable"]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td')
        for each in elements:
            head.append(each.text)
        elements = driver.find_elements_by_xpath(
            '//*[@id="testTable"]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr')
        list = []
        for each in elements:
            list.append(each.text)
        Row = len(list) - 1
    except:
        None
    tb2_list = []
    tb2_list.append(head)
    row = []
    for r in range(4, Row + 4):
        try:
            elements = driver.find_elements_by_xpath(
                '//*[@id="testTable"]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr['+str(r)+']/td')
            for each in elements:
                try:
                    ele = float(each.text)
                except:
                    ele = each.text
                row.append(ele)
        except:
                None
        tb2_list.append(row)
        row = []
    return tb2_list

#########################################################
def export_func(tb1_df,tb2_df,path):
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    # Export dataFrame as File for the specific data
    tb1_df.to_excel(writer, sheet_name='Table 1', index=None, header=None)
    tb2_df.to_excel(writer, sheet_name='Table 2', index=None, header=None)
    writer.save()
    writer.close()

#########################################################
#tobechanged
webdriver_path = 'C:\\Users\\wongh\\Downloads\\chromedriver_win32\\chromedriver'
scraping_site = 'https://www.hkab.org.hk/en/rates/exchange-rates'

chrome_options = webdriver.chrome.options.Options()
chrome_options.add_argument("--lang=en-ca")
#chrome_options.add_argument("--headless")
driver = webdriver.Chrome(webdriver_path, chrome_options = chrome_options)
driver.get(scraping_site);
wait = WebDriverWait(driver, 30)

#########################################################
#tobechanged
start_date = date(2019, 12, 30)
#tobechanged
end_date = date(2020, 1, 3)
delta = timedelta(days=1)

while start_date <= end_date:
    y = str(start_date).split('-')[0]
    m = str(start_date).split('-')[1]
    d = str(int(str(start_date).split('-')[2]))
    start_date += delta
    date_func(y, m, d)
    tb1_list = tb1_func()
    tb1_df = (pd.DataFrame(tb1_list)).dropna()
    tb2_list = tb2_func()
    tb2_df = (pd.DataFrame(tb2_list)).dropna()
    # tobechanged
    export_func(tb1_df, tb2_df, r'C:\Users\wongh\PycharmProjects\Webscraping\HKAB_ExchangeRates' + d + '.xlsx')

driver.close()