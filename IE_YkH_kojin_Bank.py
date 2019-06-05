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

ie_driver.get('https://www.boy.co.jp/kojin/benri/myd/login.html')#個人ログイン画面
time.sleep(2)
#
#テキスト入力が上手くいった
#ie_driver.execute_script("document.querySelectorAll('#ipcustomers > table > tbody > tr:nth-child(1) > td:nth-child(2)')[0].textContent = 'test';")
#IDを入力
ie_driver.execute_script("document.querySelectorAll('#rt_mypage_loginparts > div.loginpartsInner01 > form > div:nth-child(2)')[0].textContent = '7616329970';")
#PWDを入力
ie_driver.execute_script("document.querySelectorAll('#rt_mypage_loginparts > div.loginpartsInner01 > form > div:nth-child(6)')[0].textContent = 'RT10031002m';")

ie_driver.execute_script("document.querySelectorAll('#btn001')[0].click();")

