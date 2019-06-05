import time
from selenium import webdriver
import pyautogui

ie_driver = webdriver.Ie(r"C:\Users\86001\PycharmProjects\HelloTensorFlow\IEDriverServer.exe")
ie_driver.get('https://e-biz.smbc.co.jp/core/index.html')
ie_driver.maximize_window()
time.sleep(5)
#ie_driver.execute_script("document.querySelectorAll('#breadList')[0].click();")

print("AAA")

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(2218,569,0.5)
pyautogui.click(2218,569)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(1199,566,0.5)
pyautogui.click(1199,566)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(2221,587,0.5)
pyautogui.click(2221,587)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(1197,570,0.5)
pyautogui.click(1197,570)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(2236,601,0.5)
pyautogui.click(2236,601)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(1214,564,0.5)
pyautogui.click(1214,564)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(2226,623,0.5)
pyautogui.click(2226,623)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(1199,566,0.5)
pyautogui.click(1199,566)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(2214,644,0.5)
pyautogui.click(2214,644)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(1190,564,0.5)
pyautogui.click(1190,564)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(2247,1593,0.5)
pyautogui.click(2247,1593)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(1199,566,0.5)
pyautogui.click(1199,566)
time.sleep(2)

ie_driver.close()
