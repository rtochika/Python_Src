import time
from selenium import webdriver
import openpyxl

#Excelデータ操作
wb = openpyxl.load_workbook("路線data.xlsx")
#print(type(wb))
print(wb.get_sheet_names())
sheet=wb.get_sheet_by_name("Sheet1")
hatsu=sheet['B2'].value
chaku=sheet['C2'].value
print(hatsu)
print(chaku)

# chromedriverのPATHを指定（本ファイルと同じフォルダの場合）
driver_path = './chromedriver'

# Windows
# driver_path = r'./chromedriver.exe'

# Chrome起動
driver = webdriver.Chrome(driver_path)

# yahooにアクセス
driver.get('https://www.yahoo.co.jp/')

#Linkをクリック
link_elem=driver.find_element_by_partial_link_text('路線情報')
link_elem.click()

#出発を入力
nameField = driver.find_element_by_name('from')
#nameField.send_keys('元住吉')
nameField.send_keys(hatsu)
time.sleep(1)

#到着を入力
nameField = driver.find_element_by_name('to')
#nameField.send_keys('三田')
nameField.send_keys(chaku)
#time.sleep(1)

#検索ボタンをクリック
nameField = driver.find_element_by_id('searchModuleSubmit')
nameField.submit()

#Linkをクリック（最安値をクリック）
link_elem=driver.find_element_by_partial_link_text('料金の安い順')
link_elem.click()
time.sleep(3)

#料金アウトプット#
nameField = driver.find_element_by_class_name('mark')
print(nameField.text)

sheet['D2'].value=nameField.text
wb.save("路線data.xlsx")

#Chromeを終了
driver.quit()

