import time
from selenium.webdriver.common.by import By
from selenium import webdriver
from openpyxl.styles import Font
import openpyxl
import pyautogui
import pandas as pd

keyWord=input("Add keyword ")

driver=webdriver.Chrome()
driver.get("https://www.emag.ro/")
time.sleep(2)
driver.find_element(By.XPATH,'//*[@id="searchboxTrigger"]').click()
time.sleep(2)
driver.find_element(By.XPATH,'//*[@id="searchboxTrigger"]').send_keys(keyWord)
time.sleep(2)
pyautogui.press("enter")
time.sleep(2)

names=driver.find_elements(By.CLASS_NAME,'card-v2-title-wrapper')
clean_names = [element for element in names if element]

pNames=[name.text for name in clean_names]

prices=driver.find_elements(By.CLASS_NAME,'product-new-price')
pPrices=[price.text for price in prices]

ratings=driver.find_elements(By.CLASS_NAME,'star-rating-text ')
pRatings=[rating.text for rating in ratings]

print(pNames, pPrices, pRatings)

serie1 = pd.Series(pNames, name='Nume Produs')
serie2 = pd.Series(pPrices, name='Pret Produs')
serie3 = pd.Series(pRatings, name='Rating Produs')


# Combine the Series into a DataFrame
df = pd.concat([serie1, serie2, serie3], axis=1)

df.to_excel(f'{keyWord}.xlsx',index=False)

workbook=openpyxl.Workbook()
sheet=workbook.active



