import smtplib
from email.mime.text import MIMEText
from email.utils import formatdate

#-------------------------------------------------#
#mPOPメールからの送信
#-------------------------------------------------#
FROM_ADDRESS = 'auto-send@pop.keihin.co.jp'
TO_ADDRESS = 'rtochika@eg.keihin.co.jp'
BCC = ''
SUBJECT = '社内のSMTPサーバ経由'
BODY = 'pythonでメール送信'

def create_message(from_addr, to_addr, bcc_addrs, subject, body):
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Bcc'] = bcc_addrs
    msg['Date'] = formatdate()
    return msg

def send(from_addr, to_addrs, msg):
    smtpobj = smtplib.SMTP('192.168.99.7', 25)
    smtpobj.sendmail(from_addr, to_addrs, msg.as_string())
    smtpobj.close()

if __name__ == '__main__':

    to_addr = TO_ADDRESS
    subject = SUBJECT
    body = BODY

    msg = create_message(FROM_ADDRESS, to_addr, BCC, subject, body)
    send(FROM_ADDRESS, to_addr, msg)
