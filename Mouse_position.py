import pyautogui
import time
import mouse

#------------------------------------------------------------*
#マウスの右ポタンでカーソルの位置取得
#終了は一番上にマウスのカーソル持っていく
#--------------------------------------------------------------------------*
print("マウスの右ポタンでカーソルの位置を取得します。")
print("終了するにはマウスを一番上まで持っていくか、Ctrl+Cで終了します。")
print("左クリックでは何も起きません。")
#--------------------------------------------------------------------------*
loop=1
while(loop):
    time.sleep(2)
    mouse.on_right_click(lambda: print(mouse.get_position()))#右ボタンが押された？
    m_xy = mouse.get_position()
    if m_xy[1]==0:#Y軸のマウスの位置がゼロのときにブレイク
        rtn=pyautogui.confirm('終了しますか？')
        if rtn=='OK':#終了します
            break
    #print(mouse.get_position()[1])
