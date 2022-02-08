# Script gets a public google sheet as input and validates rows of 
# addresses from the remote website and result is added to the same 
# csv file.

__author__ = 'Bahram Ziaee'

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import Select

import pandas as pd
import requests
import csv
from time import sleep


# installed geckodriver version: 0.30.0


sheet_url = 'https://docs.google.com/spreadsheets/d/1H1a9eBamflt3w-4BPEk1kJYc4VgsDBWlDjkS0hV5tAY/edit#gid=0'
url = sheet_url.replace('/edit#gid=', '/export?format=csv&gid=')
usps_url = 'https://tools.usps.com/zip-code-lookup.htm?byaddress'

print("scraper strted to process input data ...\n\n")
res = requests.get(url)
open('google_sheet.csv', 'wb').write(res.content)
df = pd.read_csv('google_sheet.csv')
# print(df.info)



# iterating rows in dataframe and checking each address in a row
for index, row in df.iterrows():
    try: 
        print("checking the address in row " + str(index+1), end=" ... ")
        options = Options()
        options.headless = True
        driver = webdriver.Firefox(options=options)
        driver = webdriver.Firefox()
        driver.get(usps_url)

        driver.find_element(By.XPATH, '//*[@id="tCompany"]').send_keys(row["Company"])
        driver.find_element(By.XPATH, '//*[@id="tAddress"]').send_keys(row["Street"])
        driver.find_element(By.XPATH, '//*[@id="tCity"]').send_keys(row["City"])

        states = driver.find_element(By.XPATH, '//*[@id="tState"]')
        select = Select(states)
        select.select_by_value(row["St"])

        states = driver.find_element(By.XPATH, '//*[@id="tState"]')
        for option in states.find_elements(By.TAG_NAME, 'option'):
            if option.text == row["St"]:
                option.select()
        
        driver.find_element(By.XPATH, '//*[@id="tZip-byaddress"]').send_keys(row["ZIPCode"])
        driver.find_element(By.XPATH, '//*[@id="zip-by-address"]').click()
        
        error_elements = driver.find_elements(By.CLASS_NAME, 'server-error')
        err_list = []
        for el in error_elements:
            if el.is_displayed():
                err_list.append(el)

        if len(err_list) != 0:
            df.at[index, "isValid"] = "no"
        else:
            df.at[index, "isValid"] = "yes"

        sleep(1)
        print("completed.")
        driver.quit()
    
    except:
        print("Checking failed for address in row: " + str(index+1))


pd.DataFrame.to_csv(df, "google_sheet.csv")
print("Done.\n\n")

with open('google_sheet.csv', 'r') as file:
    reader = csv.reader(file)
    for row in reader:
        print(row)

