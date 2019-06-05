import time
from selenium import webdriver
import openpyxl

# Selectタグが扱えるエレメントに変化させる為の関数を呼び出す
from selenium.webdriver.support.ui import Select

deb_flg =0 # 1:ON 0:OFF
#【機能】---------------------------------------------------*
#収支実績表
#-----------------------------------------------------------*
glbl_flg = 0#

DCK_PL_MONTH = '2019/04/01'

# chromedriverのPATHを指定（本ファイルと同じフォルダの場合）
driver_path = './chromedriver'
# Chrome起動
driver = webdriver.Chrome(driver_path)
driver.get('http://192.168.20.211/yojitsu/(S(ttclfd55yjcu5g45kpkvb355))/menu/TopPage.aspx?Mode=6&Date=2017/06/01')

#国内
kokunai=driver.find_element_by_id('ctl00_ContentPlaceHolder1_TopDsp2_HyperLink')
kokunai.click()
time.sleep(2)

#ダックシステム
dck=driver.find_element_by_id('ctl00_ContentPlaceHolder1_内訳Syusi12_LinkButton')
dck.click()

#select option(オプションボックス)を丸一日かかって解決した！

optn_element=driver.find_element_by_id('ctl00_ContentPlaceHolder1_対象年月Syusi_DropDownList')
select_element = Select(optn_element)
#mnth=input("対象月入れて(例：2019/03/01)→ ")
#select_element.select_by_value(str(mnth))
#select_element.select_by_value('2019/04/01')#これが正しい！
select_element.select_by_value(DCK_PL_MONTH)#これが正しい！
#select_element.select_by_value('2018年04月')ではない！！！！！！！！！

time.sleep(2)

#全体
zentai=driver.find_elements_by_css_selector('#ctl00_ContentPlaceHolder1_SyusiMeisai_GridView > tbody > tr:nth-child(51)')
rp_z=zentai[0].text
rp_z=rp_z.replace(" ","/")
print("全 体",rp_z)

#物流
butsuryu=driver.find_elements_by_css_selector('#ctl00_ContentPlaceHolder1_SyusiMeisai_GridView > tbody > tr:nth-child(155)')
rp_b=butsuryu[0].text
rp_b=rp_b.replace(" ","/")
print("物 流",rp_b)

#K-NET
knet=driver.find_elements_by_css_selector('#ctl00_ContentPlaceHolder1_SyusiMeisai_GridView > tbody > tr:nth-child(207)')
rp_k=knet[0].text
rp_k=rp_k.replace(" ","/")
print("K-NET",rp_k)

driver.close()