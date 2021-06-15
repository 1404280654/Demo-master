import smtplib,time
from email.mime.text import MIMEText
# 发送邮件相关参数
smtpsever = 'smtp.126.com'
sender = 'luruifengx@126.com'
psw = 'qq123456789'
receiver = '1312360584@qq.com'
# 编辑邮件的内容
subject = u'NBA'
body = str(time.ctime())
msg = MIMEText(body,'html','utf-8')
msg['from'] = 'luruifengx@126.com'
msg['to'] = '1312360584@qq.com'
msg['subject'] = subject

try:
    smtp = smtplib.SMTP()    #获取对象
    smtp.connect(smtpsever)#链接服务器
    smtp.login(sender, psw)
    smtp.sendmail(msg['from'], sg['to'], msg.as_string())#发送动作
    print('发送成功')
except Exception as msg:
    print(msg)