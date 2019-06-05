import time
from selenium import webdriver
import openpyxl
# Selectタグが扱えるエレメントに変化させる為の関数を呼び出す
from selenium.webdriver.support.ui import Select

deb_flg =0 # 1:ON 0:OFF
#【機能】---------------------------------------------------*
#就労管理システムを立ち上げ、時間管理が必要な職員を対象に
#日時確定がされたいないか、チェックする
#-----------------------------------------------------------*
# member_loop(mcnt)　mcnt→非役職者の行数
#-----------------------------------------------------------*
glbl_flg = 0#役職あり=1、なし=0
def member_loop(mcnt):
    if mcnt < 9:
        nameField = driver.find_element_by_id('ctl00_ContentPlaceHolder1_BukaList_GridView_ctl0'+str(mcnt+1)+'_Set_Button')
        nameField.click()
    else:
        nameField = driver.find_element_by_id('ctl00_ContentPlaceHolder1_BukaList_GridView_ctl'+str(mcnt+1)+'_Set_Button')
        nameField.click()
    for cntr in range(1, 31):  # チェック確認　日付をインクリメント
        if cntr < 9:
            shusya = driver.find_element_by_id('ctl00_ContentPlaceHolder1_TimeCardList_GridView_ctl0'+str(cntr + 1)+ '_出退出社_Label')
            if(deb_flg):
                print("時間管理あり！役職者でじゃない",shusya.text)
            nameField = driver.find_element_by_id('ctl00_ContentPlaceHolder1_TimeCardList_GridView_ctl0' + str(cntr + 1) + '_日次_Label')
            if(deb_flg):
                print(nameField.text)  # 確定、OK?、未が入る
            if nameField.text=="OK?":#　Alert mail対象
                print("Alert mail対象")
        else: #10日以降
            nameField = driver.find_element_by_id('ctl00_ContentPlaceHolder1_TimeCardList_GridView_ctl' + str(cntr + 1) + '_日次_Label')
            if nameField.text=="OK?":#　Alert mail対象
                print("Alert mail対象")
            #print(nameField.text)
        # 日時確定実行
        # nameField = driver.find_element_by_id('ctl00_ContentPlaceHolder1_TimeCardList_GridView_ctl18_日次_ImageButton')
        # nameField.click()
    else:# 終了後の処理
        print("Finish!")
#-----------------------------------------------------------*

#main_strt--------------------------------------------------*
driver_path = './chromedriver'# chromedriverのPATHを指定（本ファイルと同じフォルダの場合）
driver = webdriver.Chrome(driver_path)# Chrome起動
driver.get('http://192.168.20.211/syuro/menu/TopPage.aspx?Mode=1&Rtn=Login')#就労管理システム起動
time.sleep(2)

#遠近の所属をK-NETシステム部に設定（システム統轄部では部下がいないのでエラーになる）
syozoku=driver.find_element_by_id('ctl00_ContentPlaceHolder1_所属_DropDownList')
select_element = Select(syozoku)
select_element.select_by_value('Ｋ－ＮＥＴシステム部')#valueの値をセット！
#---------------------------------------------------------------------------------*

link_elem=driver.find_element_by_id('ctl00_ContentPlaceHolder1_lbtnList')#一覧ボタンをクリック
link_elem.click()
time.sleep(1)

t_row=len(driver.find_elements_by_css_selector("tr"))#テーブルの行数を取得
if (deb_flg):
    print("len >>>>>>>>>",t_row)
for cnt in range(1,t_row-1):#部下を順番に選択
     time.sleep(1)
     if cnt < 9:#役職情報取得
        yaku_shoku=driver.find_element_by_id('ctl00_ContentPlaceHolder1_BukaList_GridView_ctl0'+str(cnt + 1)+ '_役職名_Label')
     else:
         yaku_shoku = driver.find_element_by_id('ctl00_ContentPlaceHolder1_BukaList_GridView_ctl' + str(cnt + 1) + '_役職名_Label')

     if (yaku_shoku.text):#役職者 true
         glbl_flg = 1
     else:#役職者なしのみ対象
         glbl_flg = 0

         member_loop(cnt)#関数Call
         link_elem = driver.find_element_by_id('ctl00_ContentPlaceHolder1_lbtnList')#一覧をクリック （一覧＜部下の選択画面＞に戻る）
         link_elem.click()

#main_end----------------------------------------------------*

