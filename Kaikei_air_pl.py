import csv
import openpyxl
#-------------------------------------------------#
#ケイヒン航空の５事業所の会計システムから出力した
# PLデータ（CSV）を転記する
#-------------------------------------------------#
deb_flg =1 # 1:ON 0:OFF
#------------------------------------------------------#
#関東・関西営業
#------------------------------------------------------#
def exec_sales(cnt):#関東・関西営業
    if cnt==0:#関東営業
        FILE_in_sales = r'C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/2019年04月分_902010_9021営業課_6桁損益計算書FILE.csv'
    else:#関西営業
        FILE_in_sales=r'C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/2019年04月分_903018_9031関西営業課_6桁損益計算書FILE.csv'

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
    AC_AIR_32126 = "32126" #航空輸出費用  通関費
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
    air_32126 = 0  # 航空輸出費用  通関費"
    air_32127 = 0  # 航空輸出費用  クーリエ費用(OBC)
    air_32128 = 0  # 航空輸出費用  その他"

    ACCOUNT_POS=0 #勘定科目の列番目
    TOGETSU_POS=4#当月金額の列番目

    #転送先セル-------------------------------------------------*
    #収入
    if cnt==0:#関東営業
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
    else:#関西営業
        KONSAI = "Z7"  #混載
        TESURY = "Z8"  #手数料
        NYUSYU = "Z9"  #入出庫
        UNSORO = "Z10" #運送料
        SEAAIR = "Z11" #SEA & AIR
        TSUKAN = "Z12" #通関料
        OBCSYU = "Z13" #OBC収入
        SONOTA = "Z14" #その他
        #費用
        KONGEN = "Z16" #混載原価
        NYUSYH = "Z17" #入出庫費
        UNSOHI = "Z18" #運送費
        SEAARH = "Z19" #SEA & AIR
        TSUKAH = "Z20" #通関費
        OBCHIY = "Z21" #OBC費用
        SONOTH = "Z22" #その他費用

    #----------------------------------------------------*
    cntr=0
    with open(FILE_in_sales) as f:
        #next(f)#１行目がブランクだとエラーになるので苦肉の策（２行目から読む）
        reader = csv.reader(f)
        for row in reader:
            cntr=cntr+1
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
            elif(AC_AIR_32125 in row[ACCOUNT_POS]):#航空輸出費用  情報処理費  要確認
                air_32125=air_32125+ int(row[TOGETSU_POS]) #当月金額
            elif(AC_AIR_32126 in row[ACCOUNT_POS]):#航空輸出費用  通関費
                air_32126=air_32126+ int(row[TOGETSU_POS]) #当月金額
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
    air_32126 = round(air_32126/1000)
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
    FILE_OUT = r'C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/2019年度(第31期)営業収支実績.xlsx'
    wb2 = openpyxl.load_workbook(FILE_OUT)
    print(wb2.get_sheet_names())
    print("*------2019年度(第31期)営業収支実績.xlsx------*")
    sheet2=wb2.get_sheet_by_name("4月")

    sheet2[KONSAI]=air_32111
    sheet2[TESURY]=air_32112
    sheet2[NYUSYU]=air_32113
    sheet2[UNSORO]=air_32114
    sheet2[TSUKAN]=air_32116
    sheet2[OBCSYU]=air_32117
    sheet2[SONOTA]=air_32119
    sheet2[KONGEN]=air_32121
    sheet2[NYUSYH]=air_32123
    sheet2[UNSOHI]=air_32124
    sheet2[TSUKAH]=air_32126
    sheet2[OBCHIY]=air_32127
    sheet2[SONOTH]=air_32128+air_32125
    #営業----------------------------------------------------*

    wb2.save(FILE_OUT) #ファイルへの書き込み
#------------------------------------------------------#
#成田・羽田・関空
#------------------------------------------------------#
def exec_eigyosho(cnt):#成田・羽田・関空
    if cnt==0:#成田
        FILE_eigyosho = r'C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/2019年04月分_902016_9025成田営業所_6桁損益計算書FILE.csv'
    elif cnt==1:#羽田
        FILE_eigyosho = r'C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/2019年04月分_902018_9024羽田空港営業所_6桁損益計算書FILE.csv'
    elif cnt==2:#関空
        FILE_eigyosho = r'C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/2019年04月分_903019_9034関西空港営業所_6桁損益計算書FILE.csv'

    AC_AIR_322111 = "322111" #混載手数料その１
    AC_AIR_322112 = "322112" #混載手数料その２
    AC_AIR_322114 = "322114" #混載手数料その３
    AC_AIR_322114 = "322114" #混載手数料その４
    AC_AIR_322118 = "322118" #着払い運賃
    AC_AIR_322130 = "322130" #入出庫料
    AC_AIR_322140 = "322140" #運送料
    AC_AIR_322160 = "322160" #通関料
    AC_AIR_322190 = "322190" #その他
    AC_AIR_322230 = "322230" #出庫費
    AC_AIR_322241 = "322241" #運送費
    AC_AIR_322242 = "322242" #運送傭車費
    AC_AIR_322250 = "322250" #情報処理費
    AC_AIR_322260 = "322260" #通関費
    AC_AIR_322289 = "322289" #その他

    #値の初期化
    air_322111 = 0 #混載手数料その１
    air_322112 = 0 #混載手数料その２
    air_322114 = 0 #混載手数料その３
    air_322118 = 0 #着払い運賃
    air_322130 = 0 #入出庫料
    air_322140 = 0 #運送料
    air_322160 = 0 #通関料
    air_322190 = 0 #その他

    air_322230 = 0 #出庫費
    air_322241 = 0 #運送費
    air_322242 = 0 #運送傭車費
    air_322250 = 0 #情報処理費
    air_322260 = 0 #通関費
    air_322289 = 0 #その他

    ACCOUNT_POS=0 #勘定科目の列番目
    TOGETSU_POS=4#当月金額の列番目

    #転送先セル-------------------------------------------------*
    #収入
    if cnt==0:#成田
        CC_FEE = "K26" #ＣＣＦＥＥ
        COMITT = "K28" #コミッション
        TESURY = "K29" #手数料
        NYUSYU = "K30" #入出庫
        UNSORO = "K31" #運送料
        TSUKAN = "K32" #通関料
        SONOTA = "K34" #その他
       #費用
        NYUSYH = "K36" #入出庫費
        UNSOHI = "K37" #運送費
        TSUKAH = "K38" #通関費
        SONOTH = "K40" #その他費用
    elif cnt==1:#羽田
        CC_FEE = "P26"  # ＣＣＦＥＥ
        COMITT = "P28"  # コミッション
        TESURY = "P29"  # 手数料
        NYUSYU = "P30"  # 入出庫
        UNSORO = "P31"  # 運送料
        TSUKAN = "P32"  # 通関料
        SONOTA = "P34"  # その他
        # 費用
        NYUSYH = "P36"  # 入出庫費
        UNSOHI = "P37"  # 運送費
        TSUKAH = "P38"  # 通関費
        SONOTH = "P40"  # その他費用
    elif cnt == 2:# 関空
        CC_FEE = "U26"  # ＣＣＦＥＥ
        COMITT = "U28"  # コミッション
        TESURY = "U29"  # 手数料
        NYUSYU = "U30"  # 入出庫
        UNSORO = "U31"  # 運送料
        TSUKAN = "U32"  # 通関料
        SONOTA = "U34"  # その他
        # 費用
        NYUSYH = "U36"  # 入出庫費
        UNSOHI = "U37"  # 運送費
        TSUKAH = "U38"  # 通関費
        SONOTH = "U40"  # その他費用
    #----------------------------------------------------*
    cntr=0
    with open(FILE_eigyosho) as f:
        #next(f)#１行目がブランクだとエラーになるので苦肉の策（２行目から読む）
        reader = csv.reader(f)
        for row in reader:
            cntr=cntr+1
            if(AC_AIR_322111 in row[ACCOUNT_POS]):#航空輸出収益  混載その１
                air_322111=air_322111+ int(row[TOGETSU_POS]) #当月金額
            elif(AC_AIR_322112 in row[ACCOUNT_POS]):#航空輸出収益 混載その２
                air_322112=air_322112+ int(row[TOGETSU_POS]) #当月金額
            elif(AC_AIR_322114 in row[ACCOUNT_POS]):#航空輸出収益  混載その３
                air_322114=air_322114+ int(row[TOGETSU_POS]) #当月金額
            elif(AC_AIR_322118 in row[ACCOUNT_POS]):#航空輸出収益  着払い運賃
                air_322118=air_322118+ int(row[TOGETSU_POS]) #当月金額
            elif(AC_AIR_322130 in row[ACCOUNT_POS]):#航空輸出収益  入出庫料
                air_322130=air_322130+ int(row[TOGETSU_POS]) #当月金額
            elif(AC_AIR_322140 in row[ACCOUNT_POS]):#航空輸出収益  運送料
                air_322140=air_322140+ int(row[TOGETSU_POS]) #当月金額
            elif(AC_AIR_322160 in row[ACCOUNT_POS]):#航空輸出収益  通関料
                air_322160=air_322160+ int(row[TOGETSU_POS]) #当月金額
            elif(AC_AIR_322190 in row[ACCOUNT_POS]):#航空輸出収益  その他
                air_322190=air_322190+ int(row[TOGETSU_POS]) #当月金額
         #費用
            elif(AC_AIR_322230 in row[ACCOUNT_POS]):#航空輸出収益  出庫費
                air_322230=air_322230+ int(row[TOGETSU_POS]) #当月金額
            elif(AC_AIR_322241 in row[ACCOUNT_POS]):#航空輸出収益  運送費
                air_322241=air_322241+ int(row[TOGETSU_POS]) #当月金額
            elif(AC_AIR_322242 in row[ACCOUNT_POS]):#航空輸出収益  傭車費
                air_322242=air_322242+ int(row[TOGETSU_POS]) #当月金額
            elif(AC_AIR_322250 in row[ACCOUNT_POS]):#航空輸出収益  情報処理
                air_322250=air_322250+ int(row[TOGETSU_POS]) #当月金額
            elif(AC_AIR_322260 in row[ACCOUNT_POS]):#航空輸出収益  通関費
                air_322260=air_322260+ int(row[TOGETSU_POS]) #当月金額
            elif(AC_AIR_322289 in row[ACCOUNT_POS]):#航空輸出収益  その他
                air_322289=air_322289+ int(row[TOGETSU_POS]) #当月金額

    air_322111 = round(air_322111/1000)
    air_322112 = round(air_322112/1000)
   #air_322114 = round(air_322114/1000)
   #air_322118 = round(air_322118/1000)
    air_322130 = round(air_322130/1000)
    air_322140 = round(air_322140/1000)
    air_322160 = round(air_322160/1000)
    air_322190 = round(air_322190/1000)

    air_322230 = round(air_322230/1000)
   #air_322241 = round(air_322241/1000)
   #air_322242 = round(air_322242/1000)
   #air_322250 = round(air_322250/1000)
    air_322260 = round(air_322260/1000)
   #air_322289 = round(air_322289/1000)

    ccfee = round((air_322114 + air_322118)/1000)
    itk_unso = round((air_322241+air_322242)/1000)
    sonota = round((air_322250 + air_322289)/1000)

    #------------------------------------------------------------------------
    FILE_OUT = r'C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/2019年度(第31期)営業収支実績.xlsx'
    wb2 = openpyxl.load_workbook(FILE_OUT)
    #print(wb2.get_sheet_names())
    print("*------2019年度(第31期)営業収支実績.xlsx------*")
    sheet2=wb2.get_sheet_by_name("4月")

    sheet2[CC_FEE] = ccfee
    sheet2[COMITT]=air_322112
    sheet2[TESURY]=air_322111
    sheet2[NYUSYU]=air_322130
    sheet2[UNSORO]=air_322140
    sheet2[TSUKAN]=air_322160
    sheet2[SONOTA]=air_322190
#費用
    sheet2[NYUSYH]=air_322230
    sheet2[UNSOHI]=itk_unso
    sheet2[TSUKAH]=air_322260
    sheet2[SONOTH]=sonota
    wb2.save(FILE_OUT) #ファイルへの書き込み
#------------------------------------------------------#
#管理費配賦・業績評価
#------------------------------------------------------#
def kanrihi_gyoseki_hyoka():
    #管理費配賦
    Haifu_F='C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/配賦基準（2019年度）.xlsx'
    wb_knr = openpyxl.load_workbook(Haifu_F, data_only=True)
    #業績評価
    Haifu_G='C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/ケイヒン航空4月分 営業収支(業績評価用)予算実績比較表.xlsx'
    wb_gyoseki = openpyxl.load_workbook(Haifu_G, data_only=True)
    #
    FILE_OUT = r'C:/Users/86001/PycharmProjects/HelloTensorFlow/2019年4月航空/2019年度(第31期)営業収支実績.xlsx'
    wb2 = openpyxl.load_workbook(FILE_OUT)
    print("*------営業収支実績------*")
    sheet2=wb2.get_sheet_by_name("4月")

    print("*------配賦セット！------*")
    sheet_wk = wb_knr.get_sheet_by_name("4月")  #管理費配賦表
    sheet2['F48'] = sheet_wk['G8'].value#関東営業
    sheet2['P48'] = sheet_wk['G9'].value#羽田
    sheet2['K48'] = sheet_wk['G10'].value#成田
    sheet2['Z48'] = sheet_wk['G12'].value#関西営業
    sheet2['U48'] = sheet_wk['G13'].value#関空
    #sheet2['F48']=kanri_kant_e#関東営業
    #sheet2['P48'] = kanri_haneda#羽田
    #sheet2['K48'] = kanri_narita#成田
    #sheet2['Z48'] = kanri_kans_e#関西営業
    #sheet2['U48'] = kanri_kankuu#関空

    print("*------業績セット！------*")
    sheet_wk = wb_gyoseki.get_sheet_by_name("Sheet1")  #業績表
    sheet2['F46'] = round(sheet_wk['P53'].value/1000)#関東営業
    sheet2['K46'] = round(sheet_wk['Z53'].value/1000)#成田
    sheet2['P46'] = round(sheet_wk['AE53'].value/1000)#羽田
    sheet2['Z46'] = round(sheet_wk['AJ53'].value/1000)#関西営業
    sheet2['U46'] = round(sheet_wk['AO53'].value/1000)#関空

    wb2.save(FILE_OUT) #ファイルへの書き込み
    #print("関東営業：",kanri_kant_e)
    #print("羽田　　：",kanri_haneda)
    #print("成田　　：",kanri_narita)
    #print("関西営業：",kanri_kans_e)
    #print("関空　　：",kanri_kankuu)

#------------------------------------------------------#
#メインルーチン
#------------------------------------------------------#
if __name__ == '__main__':
    for cnt in range(2):
        exec_sales(cnt)#関東・関西営業PLセット
    for cnt in range(3):
        exec_eigyosho(cnt)#成田・羽田・関空PLセット
    kanrihi_gyoseki_hyoka()#管理費配賦・業績評価