import pandas as pd
#----------------------------------------------------------#
#
#----------------------------------------------------------#

debug_flg=0 #デバッグフラグ
list_nam=[] #カナ名称
list_mny = []#金額
list_nam_mny = []#要衣装＋金額
cnt_max=0 #リストの最大数件
go_kei=0  #合計金額

#-----------------------------------------#
# 初期処理
#-----------------------------------------#
def initial():
    global list_nam
    global list_mny
    global list_nam_mny
    global go_kei
    global cnt_max

    df = pd.read_csv("C:/k-net\資金送金/伝票検索0626.csv", encoding="cp932")
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
    b4cnt = 0  # 売りげ金額トップ４のカウンター
    for cnt in range(cnt_max):
        for cnt2 in range(cnt_max):
            if (str(list_mny[cnt_max - (cnt + 1)]) in list_nam_mny[cnt2]):
                b4cnt = b4cnt + 1
                if b4cnt <= 4:
                    print("BEST", b4cnt, " ", list_nam[cnt2], " ", list_mny[cnt_max - (cnt + 1)])
                else:
                    print("その他: ", list_nam[cnt2], " ", list_mny[cnt_max - (cnt + 1)])
    print("合計：", go_kei)# 合計金額

#-----------------------------------------#
# メイン関数
#-----------------------------------------#
if __name__ == '__main__':
    initial()#
    top_4_sonota()#
    #資金送金日（2営業日後）を取得。土日・祝日なら翌営業日にする
