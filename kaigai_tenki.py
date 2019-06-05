import time
from selenium import webdriver
import openpyxl

deb_flg =0 # 1:ON 0:OFF
#【機能】---------------------------------------------------*
#
#
#-----------------------------------------------------------*

from selenium.webdriver.remote.webelement import WebElement

#シンガポールExcelデータ操作
wb1 = openpyxl.load_workbook("【C】KMTS2018.12.xlsx", data_only=True)
print(wb1.get_sheet_names())
sheet1=wb1.get_sheet_by_name("PL")#シンガポールPL（インプット）

#閲覧文書終始実績Excelデータ操作
wb2 = openpyxl.load_workbook("海外18012.xlsx")
print(wb2.get_sheet_names())
sheet2=wb2.get_sheet_by_name("S")#シンガポールPL（アウトプット）

#シンガポール
for cnt in range(9,58+1):
    sheet2['I'+str(cnt)]=sheet1['I'+str(cnt)].value
    sheet2['K'+str(cnt)]=sheet1['J'+str(cnt)].value
    sheet2['M'+str(cnt)]=sheet1['K'+str(cnt)].value
    sheet2['O'+str(cnt)]=sheet1['L'+str(cnt)].value
    sheet2['Q'+str(cnt)]=sheet1['M'+str(cnt)].value
    sheet2['S'+str(cnt)]=sheet1['N'+str(cnt)].value

wb2.save('海外18012.xlsx')

