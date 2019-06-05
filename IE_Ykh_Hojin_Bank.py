import time
# webdriver の情報
from selenium import webdriver
# html の タブの情報を取得
from selenium.webdriver.common.by import By
# キーボードを叩いた時に web ブラウザに情報を送信する
from selenium.webdriver.common.keys import Keys
# 次にクリックしたページがどんな状態になっているかチェックする
from selenium.webdriver.support import expected_conditions as EC
# 待機時間を設定
from selenium.webdriver.support.ui import WebDriverWait
# 確認ダイアログ制御
from selenium.webdriver.common.alert import Alert

ie_driver = webdriver.Ie(r"C:/Users/86001/PycharmProjects/HelloTensorFlow/IEDriverServer.exe")

ie_driver.get('https://www.bizsol.anser.ne.jp/0138c/rblgi01/I1RBLGI01-S01.do')#法人ログイン画面
ie_driver.maximize_window()
time.sleep(2)
#
#テキスト入力が上手くいった
#ie_driver.execute_script("document.querySelectorAll('#ipcustomers > table > tbody > tr:nth-child(1) > td:nth-child(2)')[0].textContent = 'test';")
#電子認証部分　2019/05/11（土）１４時半時点ではここまでとする。
#ie_driver.execute_script("document.querySelectorAll('#dscustomers > table > tbody > tr > td > input[type=image]')[0].click();")

