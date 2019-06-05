import csv
import openpyxl
#-------------------------------------------------#
#
#
#-------------------------------------------------#
deb_flg =1 # 1:ON 0:OFF

FILE_1 = 'C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/2019年04月分_902010_9021営業課_6桁損益計算書FILE.csv'
FILE_2 = 'C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/2019年04月分_902016_9025成田営業所_6桁損益計算書FILE.csv'
FILE_3 = 'C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/2019年04月分_902018_9024羽田空港営業所_6桁損益計算書FILE.csv'
FILE_4 = 'C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/2019年04月分_903019_9034関西空港営業所_6桁損益計算書FILE.csv'
FILE_5 = 'C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/2019年04月分_903018_9031関西営業課_6桁損益計算書FILE.csv'

AC_AIR_32111 = "32111" #航空輸出収益  混載
AC_AIR_32112 = "32112" #航空輸出収益  一般
AC_AIR_32113 = "32113" #航空輸出収益  入出庫料
AC_AIR_32114 = "32114" #航空輸出収益  運送料
AC_AIR_32116 = "32116" #航空輸出収益  通関料
AC_AIR_32117 = "32117" #航空輸出収益  クーリエ収益(OBC)
AC_AIR_32119 = "32119" #航空輸出収益  その他

AC_AIR_32121 = "32121" #航空輸出費用  混載支払
AC_AIR_32123 = "32123" #航空輸出費用  入出庫費
AC_AIR_32124 = "32124" #航空輸出費用  委託運送費
AC_AIR_32125 = "32125" #航空輸出費用  情報処理費  要確認"
AC_AIR_32127 = "32127" #航空輸出費用  クーリエ費用(OBC)
AC_AIR_32128 = "32128" #航空輸出費用  その他"

#値の初期化
air_32111 = 0  # 航空輸出収益  混載
air_32112 = 0  # 航空輸出収益  一般
air_32113 = 0  # 航空輸出収益  入出庫料
air_32114 = 0  # 航空輸出収益  運送料
air_32116 = 0  # 航空輸出収益  通関料
air_32117 = 0  # 航空輸出収益  クーリエ収益(OBC)
air_32119 = 0  # 航空輸出収益  その他

air_32121 = 0  # 航空輸出費用  混載支払
air_32123 = 0  # 航空輸出費用  入出庫費
air_32124 = 0  # 航空輸出費用  委託運送費
air_32125 = 0  # 航空輸出費用  情報処理費  要確認"
air_32127 = 0  # 航空輸出費用  クーリエ費用(OBC)
air_32128 = 0  # 航空輸出費用  その他"

ACCOUNT_POS=0 #勘定科目の列番目
TOGETSU_POS=4#当月金額の列番目

#転送先セル-------------------------------------------------*
#収入
KONSAI = "F7"  #混載
TESURY = "F8"  #手数料
NYUSYU = "F9"  #入出庫
UNSORO = "F10" #運送料
SEAAIR = "F11" #SEA & AIR
TSUKAN = "F12" #通関料
OBCSYU = "F13" #OBC収入
SONOTA = "F14" #その他
#費用
KONGEN = "F16" #混載原価
NYUSYH = "F17" #入出庫費
UNSOHI = "F18" #運送費
SEAARH = "F19" #SEA & AIR
TSUKAH = "F20" #通関費
OBCHIY = "F21" #OBC費用
SONOTH = "F22" #その他費用
#----------------------------------------------------*

cntr=0
with open(FILE_1) as f:
    reader = csv.reader(f)
    for row in reader:
        cntr=cntr+1
        #print(cntr)
        #print("32111" in row[0])#勘定科目検索
        if(AC_AIR_32111 in row[ACCOUNT_POS]):#航空輸出収益  混載 合致する
            air_32111=air_32111+ int(row[TOGETSU_POS]) #当月金額
        elif(AC_AIR_32112 in row[ACCOUNT_POS]):#航空輸出収益  一般 合致する:
            air_32112=air_32112+ int(row[TOGETSU_POS]) #当月金額
        elif(AC_AIR_32113 in row[ACCOUNT_POS]):#航空輸出収益  入出庫料
            air_32113=air_32113+ int(row[TOGETSU_POS]) #当月金額
        elif(AC_AIR_32114 in row[ACCOUNT_POS]):#航空輸出収益  運送料
            air_32114=air_32114+ int(row[TOGETSU_POS]) #当月金額
        elif(AC_AIR_32116 in row[ACCOUNT_POS]):#航空輸出収益  通関料
            air_32116=air_32116+ int(row[TOGETSU_POS]) #当月金額
        elif(AC_AIR_32117 in row[ACCOUNT_POS]):#航空輸出収益  クーリエ収益(OBC)
            air_32117=air_32117+ int(row[TOGETSU_POS]) #当月金額
        elif(AC_AIR_32119 in row[ACCOUNT_POS]):#航空輸出収益  その他
            air_32119=air_32119+ int(row[TOGETSU_POS]) #当月金額
        elif(AC_AIR_32121 in row[ACCOUNT_POS]):#航空輸出費用  混載支払
            air_32121=air_32121+ int(row[TOGETSU_POS]) #当月金額
        elif(AC_AIR_32123 in row[ACCOUNT_POS]):#航空輸出費用  入出庫費
            air_32123=air_32123+ int(row[TOGETSU_POS]) #当月金額
        elif(AC_AIR_32124 in row[ACCOUNT_POS]):#航空輸出費用  委託運送費
            air_32124=air_32124+ int(row[TOGETSU_POS]) #当月金額
        elif(AC_AIR_32125 in row[ACCOUNT_POS]):#航空輸出費用  情報処理費  要確認"
            air_32125=air_32125+ int(row[TOGETSU_POS]) #当月金額
        elif(AC_AIR_32127 in row[ACCOUNT_POS]):#航空輸出費用  クーリエ費用(OBC)
            air_32127=air_32127+ int(row[TOGETSU_POS]) #当月金額
        elif(AC_AIR_32128 in row[ACCOUNT_POS]):#航空輸出費用  その他"
            air_32128=air_32128+ int(row[TOGETSU_POS]) #当月金額

air_32111 = round(air_32111/1000)
air_32112 = round(air_32112/1000)
air_32113 = round(air_32113/1000)
air_32114 = round(air_32114/1000)
air_32116 = round(air_32116/1000)
air_32117 = round(air_32117/1000)
air_32119 = round(air_32119/1000)
incom_sum = air_32111+air_32112+air_32113+air_32114+air_32116+air_32117+air_32119#収入合計:転記不要
air_32121 = round(air_32121/1000)
air_32123 = round(air_32123/1000)
air_32124 = round(air_32124/1000)
air_32125 = round(air_32125/1000)
air_32127 = round(air_32127/1000)
air_32128 = round(air_32128/1000)
outcom_sum = air_32121+air_32123+air_32124+air_32125+air_32127+air_32128 #支出合計：転記不要

prft = incom_sum - outcom_sum #差益：転記不要

if deb_flg ==1:
    print("合計：", air_32111)
    print("合計：", air_32112)
    print("合計：", air_32113)
    print("合計：", air_32114)
    print("合計：", air_32116)
    print("合計：", air_32117)
    print("合計：", air_32119)
    print("収入合計：", incom_sum)#収入計
    print("合計：", air_32121)
    print("合計：", air_32123)
    print("合計：", air_32124)
    print("合計：", air_32125)
    print("合計：", air_32127)
    print("合計：", air_32128)
    print("支出合計：", outcom_sum)#支出計
    print("差益：", prft)#差益
#------------------------------------------------------------------------
#ExcelにOutput
#wb2 = openpyxl.load_workbook("海外18012.xlsx")
FILE_OUT = 'C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/2019年度(第31期)営業収支実績.xlsx'
wb2 = openpyxl.load_workbook(FILE_OUT)
print(wb2.get_sheet_names())
print("*------2019年度(第31期)営業収支実績.xlsx------*")
sheet2=wb2.get_sheet_by_name("4月")

sheet2[KONSAI]=air_32111
sheet2[TESURY]=air_32112
sheet2[NYUSYU]=air_32113
sheet2[UNSORO]=air_32114
#sheet2[SEAAIR]=air_XXXX なし
sheet2[TSUKAN]=air_32116
sheet2[OBCSYU]=air_32117
sheet2[SONOTA]=air_32119

sheet2[KONGEN]=air_32121
sheet2[NYUSYH]=air_32123
sheet2[UNSOHI]=air_32124
#sheet2[SEAARH]=air_XXX なし
#sheet2[TSUKAH]=air_XXX なし
sheet2[OBCHIY]=air_32127
sheet2[SONOTH]=air_32128+air_32125
#営業----------------------------------------------------*

wb2.save(FILE_OUT) #ファイルへの書き込み


