import os
import time
from selenium import webdriver
import openpyxl

deb_flg =0 # 1:ON 0:OFF
#【機能】---------------------------------------------------*
#Yahoo路線検索から、エクセルシートを読み込み、６か月の最安値の
#定期代金をエクセルに記載する
#-----------------------------------------------------------*
glbl_flg = 0#

#Excelデータ操作
from selenium.webdriver.remote.webelement import WebElement

# chromedriverのPATHを指定（本ファイルと同じフォルダの場合）
driver_path = './chromedriver'

# Chrome起動
driver = webdriver.Chrome(driver_path)
driver.get('https://transit.yahoo.co.jp/')
time.sleep(1)

#初期化
driver.find_element_by_xpath('//*[@id="sfrom"]').send_keys("")
driver.find_element_by_css_selector('#sto').send_keys("")

driver.find_element_by_xpath('//*[@id="sfrom"]').send_keys("田町")
driver.find_element_by_css_selector('#sto').send_keys("上野")
driver.find_element_by_css_selector('#searchModuleSubmit').submit()

time.sleep(1)

driver.back()

time.sleep(1)

#初期化
driver.find_element_by_xpath('//*[@id="sfrom"]').send_keys("")
driver.find_element_by_css_selector('#sto').send_keys("")

driver.find_element_by_xpath('//*[@id="sfrom"]').send_keys("田町")
driver.find_element_by_css_selector('#sto').send_keys("元住吉")
driver.find_element_by_css_selector('#searchModuleSubmit').submit()


print("AAA")