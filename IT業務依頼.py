import time
from selenium import webdriver
import openpyxl
#【機能】---------------------------------------------------*
#IT業務依頼システムを起動し、立場をシステム統轄部　課長に
#セットする＆その他
#-----------------------------------------------------------*

#main_strt--------------------------------------------------*
driver_path = './chromedriver'# chromedriverのPATHを指定（本ファイルと同じフォルダの場合）
driver = webdriver.Chrome(driver_path)# Chrome起動
driver.get('http://192.168.20.211/gyomuirai/menu/TopPage.aspx?s=1')#IT業務依頼システム起動
time.sleep(2)

#立場変更Linkをクリック
link_elem=driver.find_element_by_partial_link_text('立場変更')
link_elem.click()

# 立場（システム統轄部 課長）チェック
nameField = driver.find_element_by_id('ctl00_MainContent_Tachiba_RadioButtonList_0')
nameField.click()

# 選択ボタンクリック
nameField = driver.find_element_by_id('ctl00_MainContent_OK_Button')
nameField.click()

#1行目クリック
#nameField = driver.find_element_by_id('ctl00_MainContent_IncompleteListView_ctrl0_Lnk_Detail')
#nameField = driver.find_element_by_css_selector('#ctl00_MainContent_IncompleteListView_ctrl0_Lnk_Detail')#これでもOK
#nameField.click()




#main_end----------------------------------------------------*

