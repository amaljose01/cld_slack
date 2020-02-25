from selenium import webdriver
import pandas as pd
import numpy as np
import csv
import os
import time
import openpyxl


browser = webdriver.Chrome(executable_path='C:\Python37-32/chromedriver.exe')
url = "https://reporting.ondemand.com/sap/bc/mdrs/cdo?type=crp_l&db_id=spc-02&for_system=000000000310106221%2c_empty_or_null_&list=45&operation_type=9503861%2c9503840%2c9500970%2c9502590&period_type=m&solution_type=JAM%2cMPOS%2cRPOS%2cBIZX%2cKMS%2cJ2W%2cLMS%2cLMSV%2cWFA&sort_by=request_id%2ccrea_month%2ccustomer_name&status=INPROC&target_biz_type=ZH526%2cZH534%2cZH535%2cZH527%2cZH421%2cZH103&date=$today"
browser.get(url)
browser.set_window_size(1536, 824)

browser.find_element_by_xpath('//*[@id="CR_BTN_XLSX"]/i').click()
time.sleep(1)
browser.close()


workbook=openpyxl.load_workbook(r'C:\Users\I347708\Downloads\ServiceRequest.xlsx')
std=workbook['Info']
workbook.remove(std)
workbook.save(r'C:\Users\I347708\Downloads\ServiceRequest.xlsx')

data_default_source = pd.read_excel (r'C:\Users\I347708\Downloads\ServiceRequest.xlsx')
data_default_source.to_csv('amaljosecldslackdev.csv')
df1 = pd.DataFrame(data_default_source, columns= ['Service Request ID','Customer ID','Operation Type','Solution','Business Type','Country','Creation date time','Customer','Request Date'])
#print(df1)

search_condition = ['System/Tenant Setup']
prod_df = df1[df1['Operation Type'].str.contains('|'.join(search_condition))]

search_condition = ['Size Change license product only','Size Change license product and contract']
upsell_df = df1[df1['Operation Type'].str.contains('|'.join(search_condition))]

sfre_now = df1.pivot_table(index=['Operation Type','Solution','Customer'], aggfunc='size')
print(sfre_now)
#print(prod_df)
#print(upsell_df)

df1.to_csv('thefinal.csv')
prod_df.to_csv('prodprovisioning.csv')
upsell_df.to_csv('upsell_list.csv')
#sfre_now.to_csv('sfre_full.csv')

os.remove('ServiceRequest.xlsx')
os.remove('amaljosecldslackdev.csv')
os.remove('thefinal.csv')
os.remove('prodprovisioning.csv')
os.remove('upsell_list.csv')
#os.remove('sfre_full.csv')
