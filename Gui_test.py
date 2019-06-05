import pyautogui
import pyperclip
import time

print('画面サイズ：',pyautogui.size())
print('マウスポジション：',pyautogui.position())

#pyautogui.moveTo(1076, 1582, 2)
#pyautogui.moveTo(200, 500, 2)
pyautogui.moveTo(1040,1209, 2)
pyautogui.click(1040,1209, 1, 0.5, 'left')
time.sleep(1)

#立場（システム統轄）
pyautogui.click(1353, 616, 1, 0.5, 'left')
#選択
pyautogui.click(2783,702, 1, 0.5, 'left')
time.sleep(5)
#１行目をクリック
pyautogui.click(1682,1236, 1, 0.5, 'left')
time.sleep(2)
#申請をクリック
pyautogui.click(1594,399, 1, 0.5, 'left')


#pyautogui.typewrite("Good!", 0.5)#日本語の直接入力はできない！だからCopy and pasteを使う
#pyperclip.copy("文字を入力")
#pyautogui.hotkey('ctrl', 'v')
