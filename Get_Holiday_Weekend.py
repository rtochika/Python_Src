#------------------------------------------------------------------#
# 祝日と土日の判断
# 参考サイト：https://qiita.com/7110/items/4ece0ce9be0ce910ee90
#------------------------------------------------------------------#
import datetime
import jpholiday
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
            flg=2 #祝日
    #曜日取得
    #print(now_aft_pls.weekday()) #　=> 5
    # 祝日取得

    return flg # 0:通常営業日、1:土日、2:祝日

if __name__ == '__main__':
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
    print(sokin_date)

