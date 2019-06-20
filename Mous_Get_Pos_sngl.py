import pyautogui
import pyperclip
import time
import mouse
import subprocess
import os
import sys

#--------------------------------------------------------------------------*
print("#--------------------------------------------------------------------#")
print("　RPAのことならダックにおまかせロボ　Ver1.000")
print("　左クリックでカーソルの座標取得")
print("　右クリックでテキスト入力位置の座標取得")
print("#--------------------------------------------------------------------#")
#--------------------------------------------------------------------------*

TIME_SLEEP=2#time.sleep(s)デフォルト秒
all_py_src=""#ソースコード変数クリア

def say_hello():#左ボタンクリック
    global all_py_src
    global strt_flg
    global cntr

    if strt_flg==0:#開始画面のOKボタンが押されていない限り記録を開始しない
        return
    elif strt_flg==1:#開始画面のOKボタンが押されたがこの開始ボタンの座標は取得せす、次のクリックから開始する
        strt_flg=2
        return

    pos_xy = mouse.get_position()
    print(pos_xy[0],pos_xy[1])
    cntr=cntr+1
    prnt = "print('--------------【●●をクリックしました】-------------------')" + "#["+str(+cntr)+"]"
    moving="\n"+"pyautogui.moveTo("+str(pos_xy[0])+","+ str(pos_xy[1])+",0.5)"
    py_pos_src = prnt + moving + "\n" + "pyautogui.click(" + str(pos_xy[0]) + "," + str(pos_xy[1]) + ")" + "\n" + "time.sleep("+str(TIME_SLEEP)+")" + "\n\n"  # 改行コード含む
    all_py_src=all_py_src+py_pos_src
    pyperclip.copy(py_pos_src)

def say_hello2():#右ボタンクリック
    global all_py_src
    global strt_flg
    global cntr
    if strt_flg==0:#開始画面のOKボタンが押されていない限り記録を開始しない
        return
    elif strt_flg==1:#開始画面のOKボタンが押されたがこの開始ボタンの座標は取得せす、次のクリックから開始する
        strt_flg=2
        return

    pos_xy = mouse.get_position()
    print(pos_xy[0], pos_xy[1])
    cntr=cntr+1
    prnt = "print('--------------【●●を入力します。】-------------------')" + "#["+str(+cntr)+"]"
    moving = "\n" + "pyautogui.moveTo(" + str(pos_xy[0]) + "," + str(pos_xy[1]) + ",0.5)"
    py_pos_src = prnt + moving + "\n" + "pyautogui.click(" + str(pos_xy[0]) + "," + str(pos_xy[1]) + ")" + "\n" + "pyautogui.typewrite('*')" + "\n" + "time.sleep("+str(TIME_SLEEP)+")" + "\n\n"  # 改行コード含む
    all_py_src = all_py_src + py_pos_src
    pyperclip.copy(all_py_src)

callback=say_hello#左クリック
callback2=say_hello2#右クリック

mouse.on_click(callback)#クリックで関数呼び出し
mouse.on_right_click(callback2)#右ボタンで関数呼び出し

pyperclip.copy("")  # クリップボードクリアー
global loop
loop=1
strt_flg=0
cntr=0#クリック、入力の位置にシーケンシャルNO.をつける
pyautogui.alert('記録を開始しますか？')#プログラムストップ：OKボタンのクリックを待つ
strt_flg = 1

#---------------------------------------------------
if __name__ == '__main__':
    while(loop):
        pyautogui.alert('操作記録を開始しました!!\n終了するにはOKボタンをクリック')  # プログラムストップ：OKボタンのクリックを待つ
        loop=0
