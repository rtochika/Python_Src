import time
from selenium import webdriver
import openpyxl

#【機能】---------------------------------------------------*
#人探す君を起動し、ダックシステム/K-NETシステム部/遠近漁助
#を純にクリックし、属性データをプリント
#-----------------------------------------------------------*
#main_strt--------------------------------------------------*
driver_path = './chromedriver'# chromedriverのPATHを指定（本ファイルと同じフォルダの場合）
driver = webdriver.Chrome(driver_path)# Chrome起動
driver.get('https://www.yahoo.co.jp/')#
time.sleep(2)

dr=driver.find_element_by_partial_link_text("池袋暴走")

print(dr.text)
dr.click()
