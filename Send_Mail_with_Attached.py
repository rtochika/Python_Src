import smtplib
import sys
from email import encoders, utils
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import mimetypes


def attachment(filename):
    fd = open(filename, 'rb')
    mimetype, mimeencoding = mimetypes.guess_type(filename)
    if mimeencoding or (mimetype is None):
        mimetype = 'application/octet-stream'
    maintype, subtype = mimetype.split('/')
    if maintype == 'text':
        retval = MIMEText(fd.read(), _subtype=subtype)
    else:
        retval = MIMEBase(maintype, subtype)
        retval.set_payload(fd.read())
        encoders.encode_base64(retval)
    retval.add_header('Content-Disposition', 'attachment', filename=filename)
    fd.close()
    return retval

def create_message(fromaddr, toaddr, subject, message, files):
    #files=['a.xlsx'] → リスト型
    msg = MIMEMultipart()
    msg['To'] = toaddr
    msg['From'] = fromaddr
    msg['Subject'] = subject
    msg['Date'] = utils.formatdate(localtime=True)
    msg['Message-ID'] = utils.make_msgid()

    body = MIMEText(message, _subtype='plain')
    msg.attach(body)

    for filename in files:
        msg.attach(attachment(filename))
    return msg.as_string()

def send(fromaddr, toaddr, message):
    s = smtplib.SMTP('192.168.99.7', 25)
    s.sendmail(fromaddr, [toaddr], message)
    s.close()

if __name__ == '__main__':
    fromaddr = 'auto-send@pop.keihin.co.jp'
    toaddr =  'rtochika@eg.keihin.co.jp'
    #attch = ['d:/tmp/test.txt'] #テキストファイルの送信がうまくいかない!!
    attch=['d:/tmp/a.xlsx']#バイナリーファイルは問題なく送信が可能

    message = create_message(fromaddr, toaddr, 'test subject', 'test body', attch)

    send(fromaddr, toaddr, message)