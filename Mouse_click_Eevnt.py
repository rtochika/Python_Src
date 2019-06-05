import pyautogui
import pyperclip
import time
import mouse
import subprocess
import os
import sys

def say_hello():#左ボタンクリック
    pos_xy = mouse.get_position()
    print(pos_xy[0], pos_xy[1])

def say_hello2():#右ボタンクリック
    pos_xy = mouse.get_position()
    print(pos_xy[0], pos_xy[1])

callback=say_hello#左クリック
callback2=say_hello2#右クリック

mouse.on_click(callback)#クリックで関数呼び出し
mouse.on_right_click(callback2)#右ボタンで関数呼び出し

global loop
loop=1
pyautogui.alert('記録を開始しますか？')#プログラムストップ：OKボタンのクリックを待つ
#---------------------------------------------------

while(loop):
    pyautogui.alert('操作記録を開始しました。\n終了するにはOKボタンをクリック')  # プログラムストップ：OKボタンのクリックを待つ
    loop=0
