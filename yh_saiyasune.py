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

wb = openpyxl.load_workbook("路線data6か月定期.xlsx")
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

    # Linkをクリック（最安値をクリック）
    link_elem = driver.find_element_by_partial_link_text('料金の安い順')
    link_elem.click()
    time.sleep(3)

    teikii_Feld = driver.find_element_by_class_name('commuterPass')
    teikii_Feld.click()
    print("定期券→",teikii_Feld.text)
    #time.sleep(3)

    teikii_Feld = driver.find_elements_by_css_selector('#route01 > dl > dd.option > div.detail.commuterPass > dl > dd > ul > li:nth-child(3) > span')
    result = teikii_Feld[0].text
    result = result.replace("円","")#円を取る
    result = result.replace(",","")#,(かんま)を取る
    print(result)
    #time.sleep(3)

    sheet['D'+str(icnt)].value=result #6か月定期代の書き込み
    wb.save("路線data6か月定期.xlsx")

#Chromeを終了
    driver.quit()
