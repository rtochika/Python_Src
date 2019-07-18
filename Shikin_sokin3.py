import pandas as pd
import datetime
import jpholiday
import openpyxl
import smtplib
import csv

import smtplib
import urllib.parse
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

#----------------------------------------------------------#
#
#----------------------------------------------------------#

debug_flg=0 #デバッグフラグ
list_nam=[] #カナ名称
list_mny = []#金額
list_nam_mny = []#要衣装＋金額
cnt_max=0 #リストの最大数件
go_kei=0  #合計金額

TO_ADDRESS = ['rtochika@gmail.com','rtochika@eg.keihin.co.jp']#複数送信先

#-----------------------------------------#
# 初期処理
#-----------------------------------------#
def initial():
    global list_nam
    global list_mny
    global list_nam_mny
    global go_kei
    global cnt_max

    df = pd.read_csv("C:/k-net/資金送金/伝票検索.csv", encoding="cp932")
    # 上記CSVファイルはこんな感じ → http://bit.ly/2JfrPBg
    grouped = df.groupby('取引先カナ名')  # カナ名称でグループバイ
    cntr = 0
    for name, group in grouped: # nameでグループ名を受け取り、groupでグループの中身を受け取る
        cntr=cntr+1
        if debug_flg==1:
            print("cntr:",cntr)
            print(name,group["金額(円貨)"].sum())
        list_nam = list_nam + [name]
        kingaku=group["金額(円貨)"].sum()#金額の合計を出す
        list_mny = list_mny + [kingaku]#金額をリスト化（テーブルを１行に纏める）
        list_nam_mny = list_nam_mny + [name+str(kingaku)]#カナ名称と金額をくっつける。
        #list_nam_mny = list_nam_mny + [str(name) + str(kingaku)]  # カナ名称と金額をくっつける。

    list_mny.sort()#金額の小さい順でソート
    cnt_max=len(list_mny)#最大カウント
    go_kei=sum(list_mny)#金額の合計

    if debug_flg == 1:
        print(list_nam)
        print(list_mny)
        print(list_nam_mny)
        print(cnt_max)
#-----------------------------------------#
# TOP4と他を表示
#-----------------------------------------#
def top_4_sonota():
    list_kana=['G99','G20','G21','G22','G23']#売上ベスト４の名称を当該セルにセット（'G99'は使用しない）
    list_kingaku = ['J99', 'J20', 'J21', 'J22', 'J23']#売上ベスト４の金額を当該セルにセット（'J99'は使用しない）
    OUTPUT_FILE="c:/k-net/資金送金/資金送金依頼書.xlsx"#←会計システムから出力したときに（・書▲ .xlsx）ブランクが気をつける
    wb2 = openpyxl.load_workbook(OUTPUT_FILE)
    #print("*------資金送金依頼書------*")
    sheet2=wb2.get_sheet_by_name('資金送金依頼書（B)')

    sonota_kei=0
    b4cnt = 0  # 売りげ金額トップ４のカウンター
    for cnt in range(cnt_max):
        for cnt2 in range(cnt_max):
            if (str(list_mny[cnt_max - (cnt + 1)]) in list_nam_mny[cnt2]):
                b4cnt = b4cnt + 1
                if b4cnt <= 4:
                    #print("BEST", b4cnt, " ", list_nam[cnt2], " ", list_mny[cnt_max - (cnt + 1)])
                    sheet2[list_kana[b4cnt]]=list_nam[cnt2]
                    sheet2[list_kingaku[b4cnt]]=list_mny[cnt_max - (cnt + 1)]
                else:
                    #print("その他: ", list_nam[cnt2], " ", list_mny[cnt_max - (cnt + 1)])
                    sonota_kei=sonota_kei+list_mny[cnt_max - (cnt + 1)]
    sheet2['G24']="その他" #'その他'
    sheet2['J24'] = sonota_kei #その他の合計
    sokindat=set_sokinbi()#送金日の設定（２営業日後）#資金送金日（2営業日後）を取得。土日・祝日なら翌営業日にする
    sheet2['F12']=sokindat
    wb2.save(OUTPUT_FILE) #ファイルへの書き込み
    print("合計：", go_kei)# 合計金額
#----------------------------------------------------------#
#引数：何日か後の日数
#戻値：何日か後の年月日
#----------------------------------------------------------#
def get_now(plus_date):#現在日時よりplus_date後
    #---datetimeから現在日時を取得-------------#
    now = datetime.datetime.now() # => 今日の日付を取得 datetime.datetime(2017, 7, 1, 23, 15, 34, 309309)
    now_aft_pls=now + datetime.timedelta(days=plus_date) # plus_date日後（例えば2日後）の年月日をゲット
    #print(now_aft_pls)
    return now_aft_pls
#----------------------------------------------------------#
##引数：何日か後の年月日
#戻値：土日・祝日判定 0:通常営業日、1:土日休日、2:祝日休日
#----------------------------------------------------------#
def holidayhantei(now_aft_pls):
    yr_pls=now_aft_pls.year
    mt_pls=now_aft_pls.month
    dy_pls=now_aft_pls.day
# ---datetimeから曜日を取得 weekdayと曜日の対応  0: 月, 1: 火, 2: 水, 3: 木, 4: 金, 5: 土, 6: 日---#
    wk_aft_pls=now_aft_pls.weekday()
    flg=0 #祝日判定フラグ 0:通常営業日、1:土日休日、2:祝日休日
    if wk_aft_pls >= 5:
        flg=1 #土曜または日曜日
    else:
        hldy=jpholiday.is_holiday_name(datetime.date(yr_pls, mt_pls, dy_pls))
        if hldy != None:#Noneは文字列ではない！
            flg=2 #祝日→hldyｎ"海の記念日"とかが入る

    #print(now_aft_pls.weekday()) #　=> 5 曜日取得

    return flg # 0:通常営業日、1:土日、2:祝日
#--------------------------------------------#
#　送金日（翌２営業日後）をセット
#--------------------------------------------#
def set_sokinbi():
    plus_date=2#←ここに何日か後の日数を入れる（２営業日後）
    end_flg=0
    while end_flg == 0:
        now_aft_pls=get_now(plus_date) #puls_date
        flg=holidayhantei(now_aft_pls) #0以外は土日または祝日
        if flg == 0:#通常営業日 1:土日、2:祝日
            end_flg=1#終了
        else:
            plus_date=plus_date+1#１日ずつ加算

    now_aft_pls_yr=now_aft_pls.year  #年
    now_aft_pls_mt=now_aft_pls.month #月
    now_aft_pls_dy=now_aft_pls.day   #日

    wk=now_aft_pls.weekday()#曜日をゲット！
    if wk==0:
        yobi="(月)"
    elif wk==1:
        yobi="(火)"
    elif wk==2:
        yobi="(水)"
    elif wk==3:
        yobi="(木)"
    elif wk==4:
        yobi="(金)"
    elif wk==5:
        yobi="(土)"
    elif wk==6:
        yobi="(日)"
    else:
        yobi = "(＊)"#ありえない！

    sokin_date=str(now_aft_pls_yr)+("年")+str(now_aft_pls_mt)+("月")+str(now_aft_pls_dy)+("日")+yobi
    return sokin_date #送金日を返す（例：2019年7月5日（金））
    print(sokin_date)

#----------------------------------------------------#
#資金送金ファイルを添付して関係者にメールを送信する
#----------------------------------------------------#
def create_message(from_addr, to_addr, cc_addrs,bcc_addrs, subject, body):
    file_path = "c:/k-net/資金送金/資金送金依頼書.xlsx"
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = from_addr
    msg['To'] = to_addr
    #---
    msg['Cc'] = cc_addrs
    #---
    msg['Bcc'] = bcc_addrs
    msg['Date'] = formatdate()
    msg.attach(MIMEText(body, 'plain', 'utf-8'))
    # 添付ファイル
    att1 = MIMEText(open(file_path, 'rb').read(), 'base64', 'utf-8')
    att1.add_header('Content-Type', 'application/octet-stream')
    att1.add_header('Content-Disposition', 'attachment', filename="資金送金依頼書.xlsx")
    msg.attach(att1)
    return msg

def send(from_addr, to_addrs, msg):
    smtpobj = smtplib.SMTP('192.168.99.7', 25)
    smtpobj.sendmail(from_addr, to_addrs.split(","), msg.as_string())
    smtpobj.close()

def mail_soushin():

    from_address = 'auto-send@pop.keihin.co.jp'
    #---------------------------
    INPUT_CSV=r'C:/k-net/資金送金/config.txt'
    cntr=0
    toaddrs=""
    cc_addrs=""
    with open(INPUT_CSV) as f:
        reader = csv.reader(f)
        for row in reader:
            if row[0] !="#":
                if row[0]=="to":
                    toaddrs=toaddrs+ row[1]+','
                elif row[0]=="cc":
                    cc_addrs = cc_addrs + row[1]+','#CCは1名だけ設定可能

    to_addr = toaddrs
    subject = '資金送金依頼書送付の件'
    body = '資金送金依頼書を送付しますので手続きをお願い致します。'
    cc=cc_addrs
    bcc=''
    msg = create_message(from_address, to_addr,cc,bcc, subject, body)
    send(from_address, to_addr, msg)
#-----------------------------------------#
# メイン関数
#-----------------------------------------#
if __name__ == '__main__':
    print("実行を開始しました。")
    initial()#
    print("処理中です・・・")
    top_4_sonota()#
    print("関係者にメールを送信しました。")
    mail_soushin()# 関係者へメールを送信する
    print("終了しました。")
