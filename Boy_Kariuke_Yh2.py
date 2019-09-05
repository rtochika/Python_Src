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
time.sleep(2)

ie_driver.get('https://www.bizsol.anser.ne.jp/0138c/rblgi01/I1RBLGI01-S01.do')
time.sleep(3)
ie_driver.get('https://www.bizsol.anser.ne.jp/WUC_USR0104/rblgi01/BLGI001-BLGI001Info.do')