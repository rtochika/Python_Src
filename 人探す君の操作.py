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
driver.get('http://192.168.20.211/HitoSearch/menu/HitoSearch.aspx')#人探す君起動
time.sleep(2)
#ケイヒンクリック
clck_elem=driver.find_element_by_css_selector('#ctl00_ContentPlaceHolder1_WebKnetOg1_treen230 > img')
clck_elem.click()

clck_elem=driver.find_elements_by_css_selector('#ctl00_ContentPlaceHolder1_WebKnetOg1_treet239')
clck_elem[0].click()
time.sleep(2)#ここにスリープを入れないと下でエラーになる！

#clck_elem=driver.find_elements_by_css_selector('#ctl00_ContentPlaceHolder1_WebKnetOg1_dtlPerson1_ctl00_lnkbName1')
clck_elem=driver.find_elements_by_css_selector('#ctl00_ContentPlaceHolder1_WebKnetOg1_dtlPerson1 > tbody > tr:nth-child(1) > td > table > tbody > tr > td:nth-child(2)')
clck_elem[0].click()
time.sleep(2)#ここにスリープを入れないと下でエラーになる！

txt_ar=driver.find_element_by_css_selector('#ctl00_ContentPlaceHolder1_Data_TextBox')
print("AREA",txt_ar.text)
print("END")

