from selenium import webdriver
import pandas as pd
import numpy as np
import csv
import os
import time
import openpyxl


browser = webdriver.Chrome(executable_path='C:\Python37-32/chromedriver.exe')
url = "https://reporting.ondemand.com/sap/bc/mdrs/cdo?date=$today&db_id=spc-02&list=45&period_type=m&type=crp_l&operation_type=9500970&for_system=000000000310106221%2C_empty_or_null_&solution_type=JAM%2CMPOS%2CRPOS%2CBIZX%2CKMS%2CJ2W%2CLMS%2CLMSV%2CWFA&target_biz_type=ZH526%2CZH534%2CZH535%2CZH527%2CZH421%2CZH103&status=INPROC&sort_by=request_id%2Ccrea_month%2Ccustomer_name%22"
browser.get(url)
browser.set_window_size(1536, 824)

browser.find_element_by_xpath('//*[@id="CR_BTN_XLSX"]/i').click()
time.sleep(1)
browser.close()


workbook=openpyxl.load_workbook(r'C:\Users\I347708\Downloads\ServiceRequest.xlsx')
std=workbook.get_sheet_by_name('Info')
workbook.remove_sheet(std)
workbook.save(r'C:\Users\I347708\Downloads\ServiceRequest.xlsx')

data_default_source = pd.read_excel (r'C:\Users\I347708\Downloads\ServiceRequest.xlsx')
data_default_source.to_csv('amaljosecldslackdev.csv')
df1 = pd.DataFrame(data_default_source, columns= ['Service Request ID','Customer ID','Operation Type','Solution','Business Type','Country','Creation date time','Customer','Request Date'])

print(df1)
df1.to_csv('thefinal.csv')

os.remove('ServiceRequest.xlsx')
os.remove('amaljosecldslackdev.csv')
os.remove('thefinal.csv')
