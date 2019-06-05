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
print("　終了すると【d:/tmp/py_src_cod.txt】を出力します")
print("　既に【d:/tmp/py_src_cod.txt】がある場合は事前に削除しておいて下さい")
print("#--------------------------------------------------------------------#")
#--------------------------------------------------------------------------*

OUT_PUT_FILE="d:/tmp/py_src_cod.txt"#ソースコードの出力ファイル
OUT_PUT_FOLDER=r"d:\tmp"#raw(r)で\のシーケンスを/に変更できる
TIME_SLEEP=2#time.sleep(s)デフォルト秒

def out_src_file():
    path_w = OUT_PUT_FILE
    with open(path_w, mode='a') as f:
        imp1="import pyautogui\n"
        imp2="import pyperclip\n"
        imp3="import time\n"
        imp4="import mouse\n"
        imp5="import subprocess\n"
        imp6="import os\n"
        imp7="import sys\n\n"

        imp_all=imp1+imp2+imp3+imp4+imp5+imp6+imp7

        cmnt1="#-----------------------------------------------#\n"
        cmnt2="#機能：\n"
        cmnt3="#\n"
        cmnt4="#-----------------------------------------------#\n\n"
        cmnt_all=cmnt1+cmnt2+cmnt3+cmnt4

        start_time_delay="time.sleep(3)\n"#システムスタートまでの待ち時間

        f.write(imp_all+cmnt_all+start_time_delay+all_py_src)
        pyperclip.copy("")  # クリップボードクリアー
        pyautogui.alert('Pythonソースコード出力先 \n'+OUT_PUT_FILE)  #Pythonソースコード

        subprocess.run('explorer {}'.format(OUT_PUT_FOLDER))#終了と同時にフォルダを開く

all_py_src=""#ソースコード変数クリア

def say_hello():#左ボタンクリック
    global all_py_src
    global strt_flg

    if strt_flg==0:#開始画面のOKボタンが押されていない限り記録を開始しない
        return
    elif strt_flg==1:#開始画面のOKボタンが押されたがこの開始ボタンの座標は取得せす、次のクリックから開始する
        strt_flg=2
        return

    pos_xy = mouse.get_position()
    print(pos_xy[0],pos_xy[1])
    prnt="print('#--------------【●●をクリックしました】-------------------#')"
    moving="\n"+"pyautogui.moveTo("+str(pos_xy[0])+","+ str(pos_xy[1])+",0.5)"
    py_pos_src = prnt + moving + "\n" + "pyautogui.click(" + str(pos_xy[0]) + "," + str(pos_xy[1]) + ")" + "\n" + "time.sleep("+str(TIME_SLEEP)+")" + "\n\n"  # 改行コード含む
    all_py_src=all_py_src+py_pos_src
    pyperclip.copy(all_py_src)

def say_hello2():#右ボタンクリック
    global all_py_src
    global strt_flg

    if strt_flg==0:#開始画面のOKボタンが押されていない限り記録を開始しない
        return
    elif strt_flg==1:#開始画面のOKボタンが押されたがこの開始ボタンの座標は取得せす、次のクリックから開始する
        strt_flg=2
        return

    pos_xy = mouse.get_position()
    print(pos_xy[0], pos_xy[1])
    prnt = "print('#--------------【●●を入力します。】-------------------#')"
    moving = "\n" + "pyautogui.moveTo(" + str(pos_xy[0]) + "," + str(pos_xy[1]) + ",0.5)"
    py_pos_src = prnt + moving + "\n" + "pyautogui.click(" + str(pos_xy[0]) + "," + str(pos_xy[1]) + ")" + "\n" + "pyautogui.typewrite('*')" + "\n" + "time.sleep("+str(TIME_SLEEP)+")" + "\n\n"  # 改行コード含む
    all_py_src = all_py_src + py_pos_src
    pyperclip.copy(all_py_src)
#pyautogui.hotkey('ctrl', 'v')

callback=say_hello#左クリック
callback2=say_hello2#右クリック

mouse.on_click(callback)#クリックで関数呼び出し
mouse.on_right_click(callback2)#右ボタンで関数呼び出し

#初期処理---------------------------------------*
path=OUT_PUT_FOLDER
rtn=os.path.exists(path)#出力先のフォルダは存在するか？
if rtn==False:
    pyautogui.alert('ソースコードの出力先フォルダがありません\n先に'+path+'を作成して再実行してください')  #t出力先フォルダ再作成要求
    sys.exit()#プログラム終了！

path = OUT_PUT_FILE#出力ファイルの存在を確認し、あれば削除する
rtn=os.path.exists(path)#ファイルある？
if rtn:
    os.remove(path)#削除実行

pyperclip.copy("")  # クリップボードクリアー
global loob
loop=1
strt_flg=0
pyautogui.alert('記録を開始しますか？')#プログラムストップ：OKボタンのクリックを待つ
strt_flg = 1

#---------------------------------------------------

while(loop):
    pyautogui.alert('操作記録を開始しました。\n終了するにはOKボタンをクリック')  # プログラムストップ：OKボタンのクリックを待つ
    loop=0

out_src_file()