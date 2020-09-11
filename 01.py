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
import pandas as pd
from selenium.webdriver.support.ui import Select
# path = "Metadata -Anand.xlsm"
Sheet = "SHEET"
# wb = load_workbook(path)
img_path = r"F:\nvli\NGMA_DEL_04-09-2020_ANAND"
#
# dspace_dict = defaultdict(list)
# Sheet_Name = wb[Sheet]
# keys = False
# count = 0
# for row in Sheet_Name.iter_rows():
#     if not keys:
#         keys = [cell.value for cell in row]
#         continue
#     # print(keys)
#     Row=[cell.value for cell in row]
#     for cell_ind,cell in enumerate(Row):
#         # print(cell,cell_ind)
#         dspace_dict[keys[cell_ind]].append( cell)
#     if count == 100:
#         break
#     else:
#         count +=1
#
# # print(Row)
#
# print(dspace_dict)
# i = 7

# for dspace_instance in dspace_dict:
#     # print(dspace_instance["dc_identifier_value_1"])
#
#     if dspace_instance in ['image']:
#
#         text = "alh_ald-" + dspace_dict["dc_identifier_value_1"][i]
#         print(text)
#         image =[]
id = []
img =[]
for file in os.listdir(img_path):
    print(file)
    id.append(file)
for i in range(len(id)):
    count =0
    for files in os.listdir(img_path +"\\"+id[i]):
        if files.endswith("h.jpg"):
            count+=1
    img.append(count)


    # if file.startswith(text):
    #     img_path1 = os.path.join(img_path,file)
    #     print(img_path1)
    #     for pic in os.listdir(img_path1):
    #         # print(pic)
    #         if pic.endswith('h.jpg'):
    #             print(pic)
    #             path =r"{}".format(img_path1 + "\\"+ pic)
    #
    #                     image.append(path)

df = pd.DataFrame(data=id,columns=[0])
df.to_excel('ngma.xlsx', header=True, index=False)
# df = pd.DataFrame(data=img,columns=[0])
# df.to_excel('count.xlsx', header=True, index=False)
# df.to_excel('test.xlsx', header=True, index=False)
print(id)
# print(path)
# print(image[0])