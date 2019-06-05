import pyautogui
import pyperclip
import time
from selenium import webdriver

#【機能】---------------------------------------------------*
#収支実績の自動起動
#
#-----------------------------------------------------------*

deb_flg =1 # 1:ON 0:OFF

if deb_flg==1:
    print('画面サイズ：',pyautogui.size())
    print('マウスポジション：',pyautogui.position())

ie_driver = webdriver.Ie(r"C:/Users/86001/PycharmProjects/HelloTensorFlow/IEDriverServer.exe")

ie_driver.get('https://e-biz.smbc.co.jp/core/index.html')#法人ログイン画面
ie_driver.maximize_window()
time.sleep(5)

#exit()
#電子認証ログイン
print("*----------電子認証ログインしました----------*")
pyautogui.moveTo(619,706, 1)
pyautogui.click(619,706, 1, 0.5, 'left')
time.sleep(2)

print("*----------証明書の選択をクリックしました----------*")
#その他
pyautogui.click(337,533)
time.sleep(1)

pyautogui.click(489,669)#KEIHIN
time.sleep(1)

print("*----------OKをクリックしました----------*")
#OK
pyautogui.click(417,745)
time.sleep(1)

print("*----------許可をクリックしました----------*")
#許可
pyautogui.click(419,541)
time.sleep(2)

print("*----------パスワードを記入----------*")
# パスワード
pyautogui.click(492,417)
pyautogui.typewrite('jimsen20')#

print("*----------ログインをクリックしました----------*")
#ログイン
pyautogui.click(508,492)
time.sleep(3)

print("*----------振込・入出金照会等をクリックしました----------*")
#振込・入出金照会
pyautogui.click(131,246)
time.sleep(2)

print("*----------取引口座照会クリックしました----------*")
#取引口座照会
pyautogui.click(84,253)
time.sleep(2)

print("*----------入出金明細照会クリックしました----------*")
#入出金明細照会
pyautogui.click(248,422)
time.sleep(2)

print("*----------本日分クリックしました----------*")
#本日分
pyautogui.click(702,494)
time.sleep(2)

