import pandas as pd
#----------------------------------------------------------#
#
#----------------------------------------------------------#

debug_flg=0

#df = pd.read_csv("d:/tmp/伝票検索結果.csv",encoding="cp932")
#df = pd.read_csv("d:/tmp/伝票検索0626.csv",encoding="cp932")

df = pd.read_csv("C:/k-net\資金送金/伝票検索0626.csv",encoding="cp932")

grouped = df.groupby('取引先カナ名')

cntr = 0
list_nam=[]
list_mny = []
list_nam_mny = []
for name, group in grouped: # nameでグループ名を受け取り、groupでグループの中身を受け取る
    cntr=cntr+1
    if debug_flg==1:
        print("cntr:",cntr)
        print(name,group["金額(円貨)"].sum())
    list_nam = list_nam + [name]
    kingaku=group["金額(円貨)"].sum()
    list_mny = list_mny + [kingaku]
    list_nam_mny = list_nam_mny + [name+str(kingaku)]

list_mny.sort()#金額の小さい順でソート
cnt_max=len(list_mny)#最大カウント

if debug_flg==1:
    print(list_nam)
    print(list_mny)
    print(list_nam_mny)
    print(cnt_max)

b4cnt=0
for cnt in range(cnt_max):
    for cnt2 in range(cnt_max):
        if (str(list_mny[cnt_max - (cnt + 1)]) in list_nam_mny[cnt2]):
            b4cnt=b4cnt+1
            if b4cnt<=4:
                print("BEST4: ",list_nam[cnt2]," ",list_mny[cnt_max - (cnt + 1)])
            else:
                print("その他: ",list_mny[cnt_max - (cnt + 1)])
