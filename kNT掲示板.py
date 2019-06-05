import time
from selenium import webdriver
import openpyxl

deb_flg =0 # 1:ON 0:OFF
#【機能】---------------------------------------------------*
#掲示板システム自動起動
#日時確定がされたいないか、チェックする
#-----------------------------------------------------------*
# chromedriverのPATHを指定（本ファイルと同じフォルダの場合）
driver_path = './chromedriver'
# Chrome起動
driver = webdriver.Chrome(driver_path)
driver.get('https://knet02.keihin.co.jp/board/Common/login?tURL=https%253a%252f%252fknet02.keihin.co.jp%252fboard%252f')

#IDを入力
nameField = driver.find_element_by_name('uid')
nameField.send_keys('86001')
#time.sleep(1)

#パスワードを入力
nameField = driver.find_element_by_name('password')
nameField.send_keys('RT10031002m')

#実行ボタンをクリック
nameField = driver.find_element_by_id('loginSubmit')
nameField.submit()

link_elem=driver.find_element_by_link_text("全開封率")
link_elem.click()

time.sleep(2)
#ここをゲットするのにそうとう苦労した！
nameField = driver.find_elements_by_css_selector('#DataTables_Table_0 > thead > tr > th:nth-child(4) > div')#開封率
nameField[0].click()#クリック

# テーブル内容取得
#tableElem = driver.find_elements_by_css_selector("#DataTables_Table_0_wrapper")
trs=len(driver.find_elements_by_css_selector("tr"))
if (deb_flg):
    print('掲示板利用者件数→ ',trs)
for cnt in range(0,trs-8):#ケイヒングループの掲示板利用者：ヘッダー・フッター行８をマイナスする
    #cntField = driver.find_elements_by_css_selector('#DataTables_Table_0 > tbody > tr:nth-child(7) > td.sorting_1')#１行目クリック
    cntField = driver.find_elements_by_css_selector('#DataTables_Table_0 > tbody > tr:nth-child('+str(cnt+1)+') > td.sorting_1')#
    if(cnt==707):#MAX行数-1（掲示板利用者数：2019/04/23時点）
        print("STOP")
    if (deb_flg):
        print("開封率→ ",cntField[0].text)

nameField = driver.find_elements_by_css_selector('#DataTables_Table_0 > tbody > tr:nth-child(1) > td:nth-child(1)')#１行目クリック
if (deb_flg==0):
    print("所属",nameField[0].text)
    print("最後")

