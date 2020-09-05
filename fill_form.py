# from twill.commands import *
# go('http://moirepository.nvli.in/password-login')
#
# fv("1", "login_email", "dspace@nvli ")
# fv("1", "login_password", "nvli@321")
#
# submit('0')

# import requests
# session_requests = requests.session()
# login_url = 'http://moirepository.nvli.in/password-login'
# values = {'login_email' : 'dspace@nvli',
#           'login_password' : 'nvli@321'}
#
# r = session_requests.post(login_url, data=values)
# # url = 'http://moirepository.nvli.in'
# url = 'http://moirepository.nvli.in/handle/123456789/1'
#
# # values_step = {'workspace_item_id':'61', 'step':'2', 'page':'1','jsp':'/submit/edit-metadata.jsp'}
# get_form = session_requests.get(url)
# print (get_form.content)
import autoit
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import urllib, os, urllib.request
import time
from openpyxl    import *
from collections import defaultdict
from selenium.webdriver.support.ui import Select
path = "AB_MUE_03.09.2020_upload.xlsx"
Sheet = "Sculpture"
wb = load_workbook(path)
img_path = r"F:\nvli\AB_MUE_03.09.2020_ANAND"

dspace_dict = defaultdict(list)
Sheet_Name = wb[Sheet]
keys = False
count = 0
for row in Sheet_Name.iter_rows():
    if not keys:
        keys = [cell.value for cell in row]
        continue
    # print(keys)
    Row=[cell.value for cell in row]
    for cell_ind,cell in enumerate(Row):
        # print(cell,cell_ind)
        dspace_dict[keys[cell_ind]].append( cell)
    if count ==100:
        break
    else:
        count +=1

# print(Row)

print(dspace_dict)

driver = webdriver.Chrome()
# driver1 = webdriver.Chrome()
driver.implicitly_wait(5)
#
usrName = 'anand.swaroop.nvli@gmail.com'
pssWrd = "anand@1#3"

driver.maximize_window()
driver.get("http://moirepository.nvli.in/password-login")

driver.find_element_by_name('login_email').send_keys(usrName)
driver.find_element_by_name('login_password').send_keys(pssWrd)
driver.find_element_by_name('login_submit').click()
wait = WebDriverWait(driver, 2)
# driver.implicitly_wait(5)
# time.sleep(5)

# print(driver.page_source)
link = driver.find_element_by_link_text('Home')
link.click()

#
# window_after = driver.window_handles[0]
# driver.switch_to.window(window_after)
# print(driver.page_source)
# time.sleep(5)


# print(str(Allahabad Museum))

link = driver.find_element_by_xpath("/html/body/main/div[3]/div[2]/div[1]/div/div/div/h4/a")
link.click()
# window_after = driver.window_handles[0]
# driver.switch_to.window(window_after)

# time.sleep(5)

link = driver.find_element_by_link_text('Sculpture')
link.click()
# window_after = driver.window_handles[0]
# driver.switch_to.window(window_after)



link = driver.find_element_by_name('submit').click()
window_after = driver.window_handles[0]
driver.switch_to.window(window_after)
# dspace_dict = {}

# driver1.wait = WebDriverWait(driver1,10)
driver.wait = WebDriverWait(driver,10)
failed_count=0
for i in range(0,29):
    for dspace_instance in dspace_dict:
        print(dspace_instance)
        # print(dspace_instance[:-1])
        print(dspace_dict[dspace_instance][i])

        if dspace_instance in ['dc_identifier_qualifier_1','dc_format_qualifier_1','dc_format_qualifier_2',
                               'dc_format_qualifier_3','dc_format_qualifier_4','dc_coverage_qualifier_1','dc_coverage_qualifier_2','dc_type',"dc_language_iso"]:
            try:
                select = Select(driver.find_element_by_name(dspace_instance))
            except Exception:
                select = Select(driver.find_element_by_name(dspace_instance[:-1]+str(int(dspace_instance[-1])-1)))

            try:
                select.select_by_value(dspace_dict[dspace_instance][i].lower().replace(" ",""))
            except Exception:
                select.select_by_value(dspace_dict[dspace_instance][i])
                # try:
                #     select.select_by_value(dspace_dict[dspace_instance[:-1]+str(int(dspace_instance[-1])-1)][i].lower().replace(" ", ""))
                # except Exception:
                #     select.select_by_value(dspace_dict[dspace_instance[:-1]+str(int(dspace_instance[-1])-1)][i])
                #
        elif dspace_instance in ['submit_dc_coverage_add',"submit_next1","submit_next2",'submit_dc_format_add1', 'submit_dc_format_add2',"submit_upload","submit_next3","submit_next4","submit_grant","submit",'submit_dc_format_add3']:
            try:
                link = driver.find_element_by_name(dspace_instance).click()
            except Exception:
                try:
                    link = driver.find_element_by_name(dspace_instance[:-1]).click()
                except Exception:
                    link = driver.find_element_by_xpath("/html/body/header/div/div/a/img")
                    link.click()
                    time.sleep(20)

                    link = driver.find_element_by_xpath("/html/body/main/div[3]/div[2]/div[1]/div/div/div/h4/a")
                    link.click()
                    time.sleep(20)

                    link = driver.find_element_by_link_text('Sculpture')
                    link.click()
                    time.sleep(20)

                    link = driver.find_element_by_name('submit').click()
                    window_after = driver.window_handles[0]
                    driver.switch_to.window(window_after)
                    driver.wait = WebDriverWait(driver, 20)
                    print(dspace_dict["dc_identifier_value_1"][i], 'could not be uploaded')
                    # with open("fail.txt", 'a') as f:
                    #     f.write(dspace_dict["dc_identifier_value_1"][i], 'could not be uploaded\n')
                    failed_count += 1
                    break

        elif dspace_instance in ['dc_subject_']:
            keyword = dspace_dict[dspace_instance][i].split(",")
            print(keyword)
            for j in range(1,3):
                driver.find_element_by_name(dspace_instance + str(j)).send_keys(keyword[j-1].strip())
                # if j==2:
                #     link = driver.find_element_by_name("submit_dc_subject_add").click()

        elif dspace_instance in ['image']:

            # text = "alh_ald-" + dspace_dict["dc_identifier_value_1"][i]
            print(dspace_dict[dspace_instance][i])
            image = []
            for file in os.listdir(img_path):
                if file == dspace_dict[dspace_instance][i]:
                    img_path1 = os.path.join(img_path, file)
                    print(img_path1)
                    for pic in os.listdir(img_path1):
                        # print(pic)
                        if pic.endswith('h.jpg'):
                            print(pic)
                            pic_path = r"{}".format(img_path1 + "\\" + pic)
                            image.append(pic_path)




            # # # upload = driver.wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/main/div/form[2]/div[3]/div[3]/div/a/text()")))
            #
            # # # driver.find_element_by_class_name("resumable-browse").send_keys(img_path)
            # # # upload.click()
            # upload = driver.find_element_by_xpath("/html/body/main/div/form[2]/div[3]/div[3]/div/a").click()
            # handle = "[CLASS:#32770; TITLE:Open]"
            # autoit.win_wait(handle, 200)
            # autoit.control_set_text(handle, "Edit1", pic_path)
            # autoit.control_click(handle, "Button1")
            # time.sleep(30)

            for jpg in image:
                print(jpg)
                time.sleep(3)
                upload = driver.find_element_by_xpath("/html/body/main/div/form[2]/div[3]/div[3]/div/a").click()
                time.sleep(3)
                handle = "[CLASS:#32770; TITLE:Open]"
                time.sleep(3)
                autoit.win_wait(handle, 100)
                autoit.control_set_text(handle, "Edit1", jpg)
                autoit.control_click(handle, "Button1")
                time.sleep(10)

        elif dspace_instance == None:
            # with open("success.txt", "a") as f:
            #     f.write(dspace_dict["dc_identifier_value_1"][i], 'uploaded\n')
            continue
        else:
            try:
                driver.find_element_by_name(dspace_instance).send_keys(dspace_dict[dspace_instance][i])
            except Exception:
                try:
                    driver.find_element_by_name(dspace_instance[:-1]+str(int(dspace_instance[-1])-1)).send_keys(dspace_dict[dspace_instance][i])
                except Exception:
                    pass

