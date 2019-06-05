import time
from selenium import webdriver
import openpyxl

deb_flg =0 # 1:ON 0:OFF
#【機能】---------------------------------------------------*
#Yahoo株価（ケイヒン）を開き、株価を取得する
#-----------------------------------------------------------*
glbl_flg = 0#

# chromedriverのPATHを指定（本ファイルと同じフォルダの場合）
driver_path = './chromedriver'
# Chrome起動
driver = webdriver.Chrome(driver_path)
driver.get('https://stocks.finance.yahoo.co.jp/stocks/detail/?code=9312.T')

kabuka=driver.find_element_by_css_selector('#stockinf > div.stocksDtl.clearFix > div.forAddPortfolio > table > tbody > tr > td:nth-child(3)')
print("ケイヒン株価 ",kabuka.text)
time.sleep(3)
zenjitsuhi=driver.find_element_by_css_selector('#stockinf > div.stocksDtl.clearFix > div.forAddPortfolio > table > tbody > tr > td.change > span.yjSt')
print("前日比 ",zenjitsuhi.text)
