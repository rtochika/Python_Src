import time
from selenium import webdriver

driver_path = './chromedriver'# chromedriverのPATHを指定（本ファイルと同じフォルダの場合）
driver = webdriver.Chrome(driver_path)# Chrome起動
driver.get('https://www.boy.co.jp/kojin/benri/myd/login.html')#
time.sleep(2)

driver.find_element_by_css_selector('#txtBox001').send_keys("7616329970")
driver.find_element_by_css_selector('#pswd001').send_keys("RT10031002m")
driver.find_element_by_css_selector('#btn001 > img').click()
