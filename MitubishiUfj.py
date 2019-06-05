import pyautogui
import pyperclip
import time
from selenium import webdriver

#【機能】---------------------------------------------------*
#三菱UFJの現行手続き（自分の支払い）の自動化
#・慎太郎の奨学金
#・県民共済
#-----------------------------------------------------------*

deb_flg =0 # 1:ON 0:OFF

#三菱UFJ銀行のサイトオープン
driver_path = './chromedriver'# chromedriverのPATHを指定（本ファイルと同じフォルダの場合）
driver = webdriver.Chrome(driver_path)# Chrome起動
driver.get('https://entry11.bk.mufg.jp/ibg/dfw/APLIN/loginib/login?_TRANID=AA000_001')#
driver.maximize_window()
time.sleep(2)
pyautogui.moveTo(1680,809, 1)
pyautogui.click(1680,809, 1, 0.5, 'left')
#ログイン
pyperclip.copy("13031199")
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)
pyautogui.moveTo(1724,998, 1)
pyautogui.click(1724,998, 1, 0.5, 'left')
#パスワード
pyperclip.copy("10031002")
pyautogui.hotkey('ctrl', 'v')

time.sleep(2)#この時間が大事！これがないとひっかかる
#pyautogui.click(1901,1192, 1, 0.5, 'left')


time.sleep(1)




