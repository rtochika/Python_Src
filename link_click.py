import time
from selenium import webdriver

# chromedriverのPATHを指定（本ファイルと同じフォルダの場合）
driver_path = './chromedriver'

# Windows
# driver_path = r'./chromedriver.exe'

# Chrome起動
driver = webdriver.Chrome(driver_path)

# Googleにアクセス
driver.get('https://www.yahoo.co.jp/')

link_elem=driver.find_element_by_partial_link_text('路線情報')
link_elem.click()
#print(driver.page_source)
# 3秒待つ
#time.sleep (3)

# Chromeを終了
#driver.quit()

