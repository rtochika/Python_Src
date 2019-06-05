import time
from selenium import webdriver
import openpyxl

deb_flg =0 # 1:ON 0:OFF
#【機能】---------------------------------------------------*
#Yahoo路線検索から、エクセルシートを読み込み、交通費の最安値の
#をエクセルに記載する
#-----------------------------------------------------------*
glbl_flg = 0#

from selenium.webdriver.remote.webelement import WebElement
#Excelデータ操作
wb = openpyxl.load_workbook("路線data交通費.xlsx")
icnt=1
for i in range(5):
    icnt=icnt+1
#print(type(wb))
    print(wb.get_sheet_names())
    sheet=wb.get_sheet_by_name("Sheet1")
    hatsu=sheet['B'+str(icnt)].value
    chaku=sheet['C'+str(icnt)].value
    print(hatsu)
    print(chaku)

# chromedriverのPATHを指定（本ファイルと同じフォルダの場合）
    driver_path = './chromedriver'
# Chrome起動
    driver = webdriver.Chrome(driver_path)
    driver.get('https://transit.yahoo.co.jp/')

#出発を入力
    nameField = driver.find_element_by_name('from')
#nameField.send_keys('元住吉')
    nameField.send_keys(hatsu)
    #time.sleep(1)
#到着を入力
    nameField = driver.find_element_by_name('to')
#nameField.send_keys('三田')
    nameField.send_keys(chaku)
#time.sleep(1)

#検索ボタンをクリック
    nameField: WebElement = driver.find_element_by_id('searchModuleSubmit')
    nameField.submit()
#Linkをクリック（最安値をクリック）
    link_elem=driver.find_element_by_partial_link_text('料金の安い順')
    link_elem.click()
    time.sleep(3)
#料金アウトプット#
    #nameField = driver.find_elements_by_css_selector('.mainWrp .navPriority .fare .mark')  # 間に半角スペース
    nameField = driver.find_elements_by_css_selector('#rsltlst > li:nth-child(1) > dl > dd > ul > li.fare > span')  # 間に半角スペース
    print(nameField[0].text)
    #print(type(nameField))
    sheet['D'+str(icnt)].value=nameField[0].text
    wb.save("路線data.xlsx")

#Chromeを終了
    driver.quit()
