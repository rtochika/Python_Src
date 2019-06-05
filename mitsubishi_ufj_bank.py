import os
import time
from selenium import webdriver

DRIVER_PATH = os.path.join(os.path.dirname(__file__), 'chromedriver')
MUFG_TOP_URL = 'https://entry11.bk.mufg.jp/ibg/dfw/APLIN/loginib/login?_TRANID=AA000_001'

INFORMATION_TITLE = 'お知らせ - 三菱東京ＵＦＪ銀行'

# MUFGのログイン情報
ID = "your mufg account id"
PASSWORD = "your mufg account password"

try:
    browser = webdriver.Chrome(DRIVER_PATH)
    browser.get(MUFG_TOP_URL)
    time.sleep(3)

    # ログイン
    browser.find_element_by_id('account_id').send_keys(ID)
    browser.find_element_by_id('ib_password').send_keys(PASSWORD)
    browser.find_element_by_xpath('//img[@alt="ログイン"]').click()
    time.sleep(5)

    # お知らせがある場合は既読にする
    # 複数個未読がある場合の処理がまだ試せてないので怪しい・・・
    while True:
        if not browser.title == INFORMATION_TITLE:
            break

        information = browser.find_element_by_xpath(
            '//table[@class="data"]/tbody/tr')
        information.find_element_by_name('hyouzi').click()
        time.sleep(5)

        browser.find_element_by_xpath('//img[@alt="トップページへ"]').click()
        time.sleep(5)

    # 入出金明細画面に移動
    browser.find_element_by_xpath('//img[@alt="入出金明細をみる"]').click()
    time.sleep(5)

    # 明細一覧を取得して表示
    banking_list = browser.find_elements_by_xpath(
        '//table[@class="data yen_nyushutsukin_001"]/tbody/tr')
    for banking in banking_list:
        print(banking.find_element_by_class_name('date').text)
        manages = banking.find_elements_by_class_name('manage')
        for manage in manages:
            print(manage.text)
        print(banking.find_element_by_class_name('transaction').text)
        print(banking.find_element_by_class_name('balance').text)
        print(banking.find_element_by_class_name('note').text)
        print('')

    # ログアウト
    browser.find_element_by_link_text('ログアウト').click()
    time.sleep(3)
finally:
    browser.quit()