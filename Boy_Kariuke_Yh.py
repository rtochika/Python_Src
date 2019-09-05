from selenium import webdriver
import pyautogui
import pyperclip
import time
import mouse
import subprocess
import os
import sys

#-----------------------------------------------#
# From:2019/09/03（火）
#機能：横浜事務課（横浜銀行）仮受金明細処理
#-----------------------------------------------#
ie_driver = webdriver.Ie(r"C:\Users\86001\PycharmProjects\HelloTensorFlow\IEDriverServer.exe")
ie_driver.get('https://www.boy.co.jp/hojin/eb/bsd/login.html')#yokohama-bank
ie_driver.maximize_window()#最大化
time.sleep(5)

#sys.exit()

print('#--------------【ログインをクリックしました】-------------------#')
pyautogui.moveTo(620,982,0.5)
pyautogui.click(620,982)
time.sleep(2)

print('#--------------【電子証明ログインをクリックしました】-------------------#')
pyautogui.moveTo(773,698,0.5)
pyautogui.click(773,698)
time.sleep(2)

print('#--------------【その他をクリックしました】-------------------#')
pyautogui.moveTo(839,888,0.5)
pyautogui.click(839,888)
time.sleep(2)

print('#--------------【電子証明書をクリックしました】-------------------#')
pyautogui.moveTo(849,990,0.5)
pyautogui.click(849,990)
time.sleep(2)

pyautogui.moveTo(806,556,0.5)
pyautogui.click(806,556)
pyperclip.copy('8002yk')
pyautogui.hotkey('ctrl', 'v')


print('#--------------【ログインをクリックしました】-------------------#')
pyautogui.moveTo(806,556,0.5)
pyautogui.click(806,556)
time.sleep(2)

print('#--------------【振込明細紹介をクリックしました】-------------------#')
pyautogui.moveTo(491,656,0.5)
pyautogui.click(491,656)
time.sleep(2)

print('#--------------【口座をクリックしました】-------------------#')
pyautogui.moveTo(612,228,0.5)
pyautogui.click(612,228)
time.sleep(2)

print('#--------------【紹介をクリックしました】-------------------#')
pyautogui.moveTo(596,579,0.5)
pyautogui.click(596,579)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(449,606,0.5)
pyautogui.click(449,606)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(451,957,0.5)
pyautogui.click(451,957)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(714,1052,0.5)
pyautogui.click(714,1052)
time.sleep(2)

