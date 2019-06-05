import pyautogui
import pyperclip
import time
import mouse

from selenium import webdriver
#main_strt--------------------------------------------------*
driver_path = './chromedriver'# chromedriverのPATHを指定（本ファイルと同じフォルダの場合）
driver = webdriver.Chrome(driver_path)# Chrome起動

print('#--------------【IT業務依頼システムを開きました】-------------------#')
driver.get('http://192.168.20.211/gyomuirai/menu/TopPage.aspx?s=1')#IT業務依頼システム起動
driver.maximize_window()
time.sleep(2)

print('#--------------【立場変更をクリックしました】-------------------#')
pyautogui.moveTo(1034,1224,0.5)
pyautogui.click(1034,1224)
time.sleep(2)

print('#--------------【システム統轄部をクリックしました】-------------------#')
pyautogui.moveTo(1352,638,0.5)
pyautogui.click(1352,638)
time.sleep(2)

print('#--------------【選択をクリックしました】-------------------#')
pyautogui.moveTo(2786,712,0.5)
pyautogui.click(2786,712)
time.sleep(2)




