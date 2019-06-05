from email import message
import smtplib

#smtp_host = 'smtp.gmail.com'
smtp_host = '192.168.99.7'

from_email = 'rtochika@eg.keihin.co.jp'  # 送信元のアドレス
to_email = 'rtochika@eg.keihin.co.jp'  # 送りたい先のアドレス

username = 'rtochika'  # 自分のプライベートGmailのアドレス
password = 'RT10031002duck1021'  # Gmailのパスワード

# メールの内容を作成
msg = message.EmailMessage()
msg.set_content('test mail')  # メールの本文
msg['Subject'] = 'test mail(sub)'  # 件名
msg['From'] = from_email  # メール送信元
msg['To'] = to_email  # メール送信先

# メールサーバーへアクセス
server = smtplib.SMTP(smtp_host, smtp_port)
server.ehlo()
server.starttls()
server.ehlo()
server.login(username, password)
server.send_message(msg)
server.quit()