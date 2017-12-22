# -*- coding: utf-8 -*-
"""
Created on Fri Dec 22 10:30:30 2017

@author: alex.hall
"""

from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
import pyperclip
import re
import pandas
import numpy as np


driver = webdriver.Chrome() #replace with .Firefox(), or with the browser of your choice
url = "http://www.nporadio2.nl/top2000"

driver.get(url)
element=driver.find_element_by_css_selector("body")
time.sleep(5)
element.send_keys(Keys.CONTROL+'a') #navigate to the page

element.send_keys(Keys.ARROW_DOWN)
element.send_keys(Keys.CONTROL+'c')
text=pyperclip.paste()


for j in range(1,300):
    for i in range(1,15):
        element.send_keys(Keys.ARROW_DOWN)
        i+=1
        
    element.send_keys(Keys.CONTROL+'c')
    text+=pyperclip.paste()
    j+=1

f=open('top2000.txt','w')
f.write(text)

with open('top2000.txt') as f:
    textProcessed=f.readlines()
  
    
entries=[]
for l in textProcessed:
    if( bool(re.search(r"(?<!\d)\d{4}(?!\d)",l))):
        entries.append(l)
        
entries=set(entries)
    
    
excelSheet = pandas.read_excel("TOP-2000-2017.xls",sheetname="Top 2000 - 2017")
excelSheet['Jaar']=0

for i in range(0, len(excelSheet['titel'])):
    print(i)
    for entry in entries:
        if( bool(re.search(str(excelSheet['titel'][i]),entry))):
            excelSheet.iloc[i,3]=entry[-5:]


           
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pandas.ExcelWriter('top2000_with_year.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
excelSheet.to_excel(writer, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
writer.save()

#page=browser.find_element_by_css_selector('body')
#page.send_keys(Keys.CONTROL+'a')
#page=browser.find_element_by_css_selector('body')
#page.send_keys(Keys.CONTROL+'a')