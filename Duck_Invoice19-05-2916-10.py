import pyautogui
import pyperclip
import time
import mouse
import subprocess
import os
import sys

#-----------------------------------------------#
#Duck_Invoice Auto Excec System
#機能：
#
#-----------------------------------------------#

time.sleep(3)
print('#--------------【利用システムをクリックしました】-------------------#')
pyautogui.moveTo(1121,434,0.5)
pyautogui.click(1121,434)
time.sleep(2)

print('#--------------【システム選択をクリックしました】-------------------#')
pyautogui.moveTo(1105,496,0.5)
pyautogui.click(1105,496)
time.sleep(2)

print('#--------------【選択をクリックしました】-------------------#')
pyautogui.moveTo(903,525,0.5)
pyautogui.click(903,525)
time.sleep(2)

print('#--------------【新規作成をクリックしました】-------------------#')
pyautogui.moveTo(643,269,0.5)
pyautogui.click(643,269)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(643,269,0.5)
pyautogui.click(643,269)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(666,117,0.5)
pyautogui.click(666,117)
time.sleep(2)

print('#--------------【請求年月をクリックしました】-------------------#')
pyautogui.click(894,393)
pyperclip.copy("2019/06")
pyautogui.hotkey('ctrl', 'v')
#pyautogui.typewrite("2019/06")
time.sleep(2)

print('#--------------【部署コードをクリックしました】-------------------#')
pyautogui.click(907,428)
pyautogui.typewrite("1036")
time.sleep(2)

print('#--------------【ＯＫをクリックしました】-------------------#')
pyautogui.moveTo(954,473,0.5)
pyautogui.click(954,473)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(1360,222,0.5)
pyautogui.click(1360,222)
time.sleep(3)

print('#--------------【請求日付を入力しました】-------------------#')
pyautogui.click(1366,218)
pyperclip.copy("2019/05/31")
pyautogui.hotkey('ctrl', 'v')
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(1338,170,0.5)
pyautogui.click(1338,170)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.click(995,120)
pyautogui.typewrite("103635001")
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(1409,121,0.5)
pyautogui.click(1409,121)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(1405,122,0.5)
pyautogui.click(1405,122)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(974,616,0.5)
pyautogui.click(974,616)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.click(897,387)
pyautogui.dragTo(480, 387, 2, button='left')
pyperclip.copy("恋するには遅すぎると言われる私でも、遠い昔に・・・")
pyautogui.hotkey('ctrl', 'v')
time.sleep(2)

time.sleep(3)
print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(640,174,0.5)
pyautogui.click(640,174)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(896,452,0.5)
pyautogui.click(896,452)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(920,610,0.5)
pyautogui.click(920,610)
time.sleep(2)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(964,609,0.5)
pyautogui.click(964,609)
time.sleep(4)

print('#--------------【●●をクリックしました】-------------------#')
pyautogui.moveTo(974,612,0.5)
pyautogui.click(974,612)
time.sleep(2)

