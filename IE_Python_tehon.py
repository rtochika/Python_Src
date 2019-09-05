from selenium import webdriver
import pyautogui
import pyperclip
import time
import mouse
import subprocess
import os
import sys

#-----------------------------------------------#
# From:2019/09/03（火）
#機能：Yhaoo 路線検索のInternet Exploler バージョン
#-----------------------------------------------#
ie_driver = webdriver.Ie(r"C:\Users\86001\PycharmProjects\HelloTensorFlow\IEDriverServer.exe")
ie_driver.get('https://www.yahoo.co.jp/')#yahoo top page
#ie_driver.maximize_window()#最大化
time.sleep(3)

#路線検索
#ie_driver.get('https://rdsig.yahoo.co.jp/_ylt=A7dPR3I5Nm9dbG0ATRiJBtF7/RV=2/RE=1567655865/RH=cmRzaWcueWFob28uY28uanA-/RB=Yohya9mzFxAVyUUtCKM7hzkRbjE-/RU=aHR0cHM6Ly90cmFuc2l0LnlhaG9vLmNvLmpwLwA-/RK=0/RS=PLv7p4DLl_co0ndNFoXOmLI3Ltc-')
ie_driver.get('http://bit.ly/2N2Cu6H')
time.sleep(2)
#出発
ie_driver.find_element_by_name("from").send_keys("元住吉")
ie_driver.find_element_by_name("to").send_keys("東京")

#検索ボタンクリック＜＞
ie_driver.find_element_by_id("searchModuleSubmit").submit()
#検索ボタンクリック＜以下でもOK＞
#ie_driver.execute_script("document.querySelectorAll('#searchModuleSubmit')[0].click();")
time.sleep(5)
#ie_driver.find_element_by_partial_link_text("料金の安い順").click()
ie_driver.execute_script('document.querySelector("#tabflt > li:nth-child(3) > a").click();')


time.sleep(5)
#<定期券>
ie_driver.execute_script('document.querySelector("#route01 > dl > dd.option > ul > li.commuterPass > p > a").click();')
#<６ヶ月金額>
teikii_Feld = ie_driver.find_elements_by_css_selector('#route01 > dl > dd.option > div.detail.commuterPass > dl > dd > ul > li:nth-child(3) > span')
print("text",teikii_Feld[0].text)

time.sleep(3)
