from selenium import webdriver
from selenium.webdriver.chrome.options import Options


# ブラウザーを起動
options = Options()
options.binary_location = '/Applications/Google Chrome Canary.app/Contents/MacOS/Google Chrome Canary'
options.add_argument('--headless')

driver_path = 'C:/Users/86001/PycharmProjects/HelloTensorFlow/chromedriver.exe'

driver = webdriver.Chrome(chrome_options=options)

# Google検索画面にアクセス
driver.get('https://www.google.co.jp/')

# htmlを取得・表示
html = driver.page_source
print(html)

# ブラウザーを終了
driver.quit()