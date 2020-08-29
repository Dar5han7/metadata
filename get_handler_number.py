import autoit
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import urllib, os, urllib.request
import time

from openpyxl import *
from collections import defaultdict
from selenium.webdriver.support.ui import Select


driver = webdriver.Chrome()
driver.implicitly_wait(10)
#
usrName = 'anand.swaroop.nvli@gmail.com'
pssWrd = "anand@1#3"

driver.maximize_window()
driver.get("http://moirepository.nvli.in/password-login")

driver.find_element_by_name('login_email').send_keys(usrName)
driver.find_element_by_name('login_password').send_keys(pssWrd)
driver.find_element_by_name('login_submit').click()
wait = WebDriverWait(driver, 2)
time.sleep(5)

# print(driver.page_source)
driver.find_element_by_name('submit_own').click()
time.sleep(10)

all_handler_details = {}
excel_file_name = 'handler.xlsx'
try:
    wb = load_workbook(excel_file_name)
except:
    wb = Workbook()



ws = wb.active
ws["A1"] = "Handler No"
ws["B1"] = "Accession No"
ws["C1"] = "index No"

for i in range(900,1405):
    # // *[ @ id = "content"] / div[3] / div / div[1] / table / tbody / tr[6] / td[2] / a
    driver.find_element_by_xpath('//*[@id="content"]/div[2]/table/tbody/tr['+str(i)+']/td[2]/a').click()
    # try://*[@id="content"]/div[3]/div/div[1]/div[2]/table/tbody/tr[6]/td[2]
    #     handle_number = driver.find_element_by_xpath('// *[ @ id = "content"] / div[3] / div / div[1] / table / tbody / tr[6] / td[2] / a').text
    #
    # except Exception:
    #     handle_number = driver.find_element_by_xpath('//*[@id="content"]/div[3]/div/div[1]/table/tbody/tr[5]/td[2]/a').text
    handle_number = driver.find_element_by_xpath('// *[ @ id = "content"] / div[3] / div / div[1] / div[1] / code').text


    accesion_number = driver.find_element_by_xpath('//*[@id="content"]/div[3]/div/div[1]/div[2]/table/tbody/tr[2]/td[1]/a').text
    ws[f"A{i}"] = handle_number
    ws[f"B{i}"] = accesion_number
    ws[f"C{i}"] = i
    print(i,handle_number,accesion_number)
    driver.back()

wb.save(excel_file_name)
wb.close()

