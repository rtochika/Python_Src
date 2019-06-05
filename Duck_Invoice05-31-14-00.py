import csv
import pyautogui
import pyperclip
import time
import mouse
import subprocess
import os
import sys
from datetime import datetime
#-----------------------------------------------#
#Date     ：2019/05/30（木）
#Python Version:3.6xx
#File Name：DateFormat.py
#システム統轄部請求データ出力
#-----------------------------------------------#
INPUT_CSV =r'c:\k-net\自動計上.csv'
denpyo_hiduke7=""#例）2019/05
denpyo_hiduke10=""#例）2019/05/31
kanri_bango=""
tekiyo=""

#--------------------------------------------------------------------------------------
class DateFormat():#0を埋めるクラス：2019/1/1→2019/01/01
    def __init__(self, date,flg):
        self.date=date
        self.flg=flg

    def format(self):
        lng=len(self.date)
        #print("Leng:", lng)
        if lng==8:#2019/1/1→2019/01/01
            self.date=self.date.replace("/","/0")
        elif lng==9:#2019/11/1 または 2019/1/11
            if self.date[7:8]=="/":#日付が一桁 → 2019/11/1 →2019/11/01
                self.date = self.date[0:8] + "0" + self.date[8:9]
            else:#月が一桁 → 2019/1/11 →2019/01/11
                self.date=self.date[0:5]+"0"+self.date[5:9]

        if self.flg==1:
            self.date =self.date[0:7]#例）self.date→2019/01
        #print("date_rtn", self.date)
#--------------------------------------------------------------------------------------

def int():#初期処理
    #請求システム起動
    subprocess.Popen(r'C:\INTLDKSM\INTLDKSMLogin.exe')
    time.sleep(3)

def print_invoice():#印刷処理
    cntr=0
    global denpyo_hiduke7
    global denpyo_hiduke10
    global kanri_bango
    global tekiyo
    with open(INPUT_CSV) as f:
        reader = csv.reader(f)
        for row in reader:
            if row[0]:#データあり
                if row[0]=='仮登録':
                    #denpyo_hiduke10=row[2]#伝票日付
                    dt = DateFormat(row[2], 0)
                    dt.format()
                    denpyo_hiduke10=dt.date
                    print("denpyo_hiduke10: ",denpyo_hiduke10)
                    dt = DateFormat(dt.date, 1)
                    dt.format()
                    denpyo_hiduke7 = dt.date
                    kanri_bango=row[7]#管理番号
                    tekiyo=row[9]#適用
                    cntr=cntr+1
                    output_sys_tokatsu(cntr)
                    print("伝票日付10:",denpyo_hiduke10,"伝票日付7:",denpyo_hiduke7," 管理番号:",kanri_bango," 適用",tekiyo)

def output_sys_tokatsu(cntr):
    time.sleep(3)
    print('#--------------【利用システムをクリックしました】-------------------#')
    pyautogui.moveTo(1121,434,0.5)
    pyautogui.click(1121,434)
    time.sleep(2)

    print('#--------------【システム選択をクリックしました】-------------------#')
    pyautogui.moveTo(1105,496,0.5)
    pyautogui.click(1105,496)
    time.sleep(1)

    print('#--------------【選択をクリックしました】-------------------#')
    pyautogui.moveTo(903,525,0.5)
    pyautogui.click(903,525)
    time.sleep(1)

    print('#--------------【新規作成をクリックしました】-------------------#')
    pyautogui.moveTo(643,269,0.5)
    pyautogui.click(643,269)
    time.sleep(1)

    if cntr == 1:#１回目だけ
        print('#--------------【●●をクリックしました】-------------------#')
        pyautogui.moveTo(643,269,0.5)
        pyautogui.click(643,269)
        time.sleep(1)

        print('#--------------【●●をクリックしました】-------------------#')
        pyautogui.moveTo(666,117,0.5)
        pyautogui.click(666,117)
        time.sleep(1)

        print('#--------------【請求年月をクリックしました】-------------------#')
        pyautogui.click(894,393)
        pyperclip.copy(denpyo_hiduke7)
        print("denpyo_hiduke7:",denpyo_hiduke7)
        #pyperclip.copy("2019/06")
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)

        print('#--------------【部署コードをクリックしました】-------------------#')
        pyautogui.click(907,428)
        pyautogui.typewrite("1036")#固定
        time.sleep(1)

    print('#--------------【ＯＫをクリックしました】-------------------#')
    pyautogui.moveTo(954,473,0.5)
    pyautogui.click(954,473)
    time.sleep(2)

    print('#--------------【●●をクリックしました】-------------------#')
    pyautogui.moveTo(1360,222,0.5)
    pyautogui.click(1360,222)
    time.sleep(2)

    print('#--------------【請求日付を入力しました】-------------------#')
    pyautogui.click(1366,218)
    pyperclip.copy(denpyo_hiduke10)
    print("kanri_bango:",kanri_bango)
    #pyperclip.copy("2019/05/31")
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)

    print('#--------------【●●をクリックしました】-------------------#')
    pyautogui.moveTo(1338,170,0.5)
    pyautogui.click(1338,170)
    time.sleep(1)

    print('#--------------【●●をクリックしました】-------------------#')
    pyautogui.click(995,120)
    pyautogui.typewrite(str(kanri_bango))#管理番号
    print("kanri_bango:",kanri_bango)
    #pyautogui.typewrite("103635001")
    time.sleep(1)

    print('#--------------【●●をクリックしました】-------------------#')
    pyautogui.moveTo(1409,121,0.5)
    pyautogui.click(1409,121)
    time.sleep(1)

    print('#--------------【●●をクリックしました】-------------------#')
    pyautogui.moveTo(1405,122,0.5)
    pyautogui.click(1405,122)
    time.sleep(1)

    print('#--------------【●●をクリックしました】-------------------#')
    pyautogui.moveTo(974,616,0.5)
    pyautogui.click(974,616)
    time.sleep(1)

    print('#--------------【●●をクリックしました】-------------------#')
    pyautogui.click(897,387)
    pyautogui.dragTo(480, 387, 2, button='left')
    pyperclip.copy(tekiyo)#適用
    #pyperclip.copy("恋するには遅すぎると言われる私でも、遠い昔に・・・")
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)

    print('#--------------【●●をクリックしました】-------------------#')
    pyautogui.moveTo(640,174,0.5)
    pyautogui.click(640,174)
    time.sleep(1)

    print('#--------------【●●をクリックしました】-------------------#')
    pyautogui.moveTo(896,452,0.5)
    pyautogui.click(896,452)
    time.sleep(1)

    print('#--------------【●●をクリックしました】-------------------#')
    pyautogui.moveTo(920,610,0.5)
    pyautogui.click(920,610)
    time.sleep(1)

    print('#--------------【●●をクリックしました】-------------------#')
    pyautogui.moveTo(964,609,0.5)
    pyautogui.click(964,609)
    time.sleep(2)

    print('#--------------【●●をクリックしました】-------------------#')
    pyautogui.moveTo(974,612,0.5)
    pyautogui.click(974,612)
    time.sleep(4)

#-------------------------------------------------------
#スタート
#-------------------------------------------------------
if __name__ == '__main__':
    int()#初期処理
    print_invoice()
    #output_sys_tokatsu()