import time
from selenium import webdriver
# Selectタグが扱えるエレメントに変化させる為の関数を呼び出す
from selenium.webdriver.support.ui import Select

import openpyxl
#【機能】---------------------------------------------------*
#
#
#-----------------------------------------------------------*

#main_strt--------------------------------------------------*
driver_path = './chromedriver'# chromedriverのPATHを指定（本ファイルと同じフォルダの場合）
driver = webdriver.Chrome(driver_path)# Chrome起動
driver.get('http://192.168.20.211/gyomuirai/menu/TopPage.aspx?s=1')#IT業務依頼システム起動
time.sleep(2)

#検索Linkをクリック
link_elem=driver.find_element_by_partial_link_text('検索')
link_elem.click()
time.sleep(2)

optn_element=driver.find_element_by_id('ctl00_MainContent_Drpl_UketsukeKubun')
select_element = Select(optn_element)
select_element.select_by_value('1')#Ｋ－ＮＥＴ
time.sleep(1)

#進捗オプションボタン→"開発"
driver.find_element_by_css_selector('#ctl00_MainContent_UketsukeSts_RadioButtonList_2').click()

#テーブルの件数をゲット
#trs=len(driver.find_elements_by_css_selector("tr"))
#print("K-NETの開発中件数:",trs)

#最大件数取得--------------------------------------------*
cntrdata = driver.find_element_by_id('ctl00_MainContent_SystemIraiListView_SearchListViewDataPager')
print('cntrdata.text:',cntrdata.text[12:15].strip())
mx=int(cntrdata.text[12:15].strip())#最大件数取得

cnt=0
for cnt in range(mx):
    cnt=cnt+1
    #data = driver.find_elements_by_css_selector("#tabView1 > div.tableDiv > div.listViewDiv > ul > li:nth-child(1) > table > tbody > tr:nth-child(1)")
    data = driver.find_elements_by_css_selector('#tabView1 > div.tableDiv > div.listViewDiv > ul > li:nth-child('+str(cnt)+') > table > tbody > tr:nth-child(1)')
    data1 = data[0].text
    # idx=data1.find('2019/04/16') #日付の位置抽出
    # print('idx:',idx)
    print("data：",data1[21:31])#日付データ抽出 例：2019/04/16

exit()

# 選択ボタンクリック
nameField = driver.find_element_by_id('ctl00_MainContent_OK_Button')
nameField.click()

#main_end----------------------------------------------------*
