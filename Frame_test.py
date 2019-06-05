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
#driver.get('https://entry11.bk.mufg.jp/ibg/dfw/APLIN/loginib/login?_TRANID=AA000_001')
#time.sleep(3)

#driver.get('https://direct.bk.mufg.jp/')
#time.sleep(3)
#tt=driver.find_element_by_css_selector('#lnav_direct_login > img')
#tt.click()
#time.sleep(3)

driver.get('https://entry11.bk.mufg.jp/ibg/dfw/APLIN/loginib/login?_TRANID=AA000_001')
time.sleep(2)

#dr=driver.find_element_by_id('account_id')
dr=driver.find_element_by_xpath('//*[@id="account_id"]')

dr.send_keys("13031199")
dr=driver.find_element_by_id('ib_password')
dr.send_keys("10031002")
#dr=drive.find_element_by_xpath('//img[@alt="ログイン"]')#これが正解！
#dr=drive.find_element_by_xpath('//*[@id="login_frame"]/div/div/div[2]/a/img')
#dr=dr.find_element_by_css_selector("#login_frame > div > div > div.acenter.admb_m > a > img")
dr.click()

#prtint("AAA")
