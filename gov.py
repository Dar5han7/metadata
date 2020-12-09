from selenium import webdriver
import requests
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import urllib, os, urllib.request
import time
# from openpyxl    import *
from collections import defaultdict
from selenium.webdriver.support.ui import Select
url = "http://www.mca.gov.in/mcafoportal/viewCompanyMasterData.do"
cin = "U40108HP1995PTC016183"
driver = webdriver.Chrome(ChromeDriverManager().install())
# driver1 = webdriver.Chrome()
driver.implicitly_wait(5)

driver.maximize_window()
driver.get(url)

driver.find_element_by_name("companyID").send_keys(cin)

img = driver.find_element_by_id('captcha')
from PIL import Image

def get_captcha(driver, element, path):
    # now that we have the preliminary stuff out of the way time to get that image :D
    location = element.location
    size = element.size
    print(size,location)
    # saves screenshot of entire page
    driver.save_screenshot(path)

    # uses PIL library to open image in memory
    image = Image.open(path)

    left = location['x']
    top = location['y']
    right = location['x'] + size['width']
    bottom = location['y'] + size['height']

    image = image.crop((left, top, right, bottom))  # defines crop points
    image.save(path, 'png')  # saves new cropped image
path = "captcha.png"
get_captcha(driver, img, "captcha.png")
# download the image
# urllib.request.urlretrieve(url, "captha.jpg")


#
# import pytesseract
# import sys
# import argparse
#
# try:
#     import Image
# except ImportError:
#     from PIL import Image
#
# from subprocess import check_output
#
#
# def resolve(path):
#     print("Resampling the Image.")
#     check_output(['convert', path, '-resample', '600', path])
#     return pytesseract.image_to_string(Image.open(path))
#
# argparser = argparse.ArgumentParser()
# argparser.add_argument('path', help='Captcha file path')
# args = argparser.parse_args()
# path = args.path
# print('Resolving Captcha')
# captcha_text = resolve(path)
# print('Extracted Text: ', captcha_text)

# import requests
#
# payload = {'apikey': "6324bb0c8a88957", "OCREngine": 2}
# f_path = "captcha.png"
# with open(f_path, 'rb') as f:
#     j = requests.post('https://api.ocr.space/parse/image', files={f_path: f}, data=payload).json()
#     if j['ParsedResults']:
#         print(j['ParsedResults'][0]['ParsedText'])

import pytesseract
import os
import argparse
try:
    import Image, ImageOps, ImageEnhance, imread
except ImportError:
    from PIL import Image, ImageOps, ImageEnhance

def solve_captcha(path):

    """
    Convert a captcha image into a text,
    using PyTesseract Python-wrapper for Tesseract
    Arguments:
        path (str):
            path to the image to be processed
    Return:
        'textualized' image
    """
    image = Image.open(path).convert('RGB')
    image = ImageOps.autocontrast(image)

    filename = "{}.png".format(os.getpid())
    print(filename)
    image.save(filename)

    text = pytesseract.image_to_string(Image.open(filename))
    return text


# if __name__ == '__main__':
argparser = argparse.ArgumentParser()
argparser.add_argument("-i", "--image", required=True, help="path to input image to be OCR'd")
args = vars(argparser.parse_args())
path = args["image"]
print('-- Resolving')
captcha_text = solve_captcha(path)
print('-- Result: {}'.format(captcha_text))
