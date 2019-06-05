import imaplib, re, email, six

e_mail_default_encoding = 'iso-2022-jp'

UserName = "rtochika@gmail.com"
PassWord = "RT10031002duck1021"

gmail = imaplib.IMAP4_SSL("imap.gmail.com",'993')
gmail.login(UserName,PassWord)
gmail.select("INBOX")

type, [data] = gmail.search(None,'UNSEEN')

count = 1


for num in data.split():
    print("カウント=",count)
    count+=1
    result, d = gmail.fetch(num,"(RFC822)")
    raw_email = d[0][1]

    #文字コード取得
    msg = email.message_from_string(raw_email.decode('utf-8'))
    msg_encoding = email.header.decode_header(msg.get('Subject'))[0][1] or 'iso-2022-jp'

    #パースして解析準備
    msg = email.message_from_string(raw_email.decode(msg_encoding))

    print('msg', msg)
#差出人情報を取得
    fromObj = email.header.decode_header(msg.get('From'))
    addr = ""
    for f in fromObj:
        if isinstance(f[0],bytes):
            addr += f[0].decode(msg_encoding)
        else:
            addr += f[0]
        print(addr)

#件名の取得＆表示
    subject = email.header.decode_header(msg.get('Subject'))
    title = ""
    for sub in subject:
        if isinstance(sub[0],bytes):
            title += sub[0].decode(msg_encoding)
        else:
            title += sub[0]
        print(title)