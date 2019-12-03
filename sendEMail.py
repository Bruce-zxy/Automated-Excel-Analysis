# -*- coding: utf-8 -*-
import os
import time
from email import Encoders
from email.header import Header
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.utils import parseaddr, formataddr
import smtplib
import base64


def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr(( \
        Header(name, "utf-8").encode(), \
        addr.encode("utf-8") if isinstance(addr, unicode) else addr))

from_addr = "admin@zhuxiaoyi.com"
password = "sebyrvoozedbbigd"
smtp_server = "smtp.qq.com"
to_addr1 = "460351637@qq.com"
to_addr2 = "bruce_zxy@163.com"

msg = MIMEMultipart()
msg.attach(MIMEText("汇总文件在附件，请查收！", "plain", "utf-8"))

defaultPath = os.getcwd() + '\\'
filename = '【每日销售情况汇报-'.decode('utf-8').encode('gbk') + time.strftime("%m.%d", time.localtime()) + '】.txt'.decode('utf-8').encode('gbk') 

# 添加附件就是加上一个MIMEBase，从本地读取一个图片:
with open(defaultPath + filename, 'rb') as f:
	# 设置附件的MIME和文件名，这里是png类型:
	mime = MIMEBase('application', 'octet-stream')
	# 加上必要的头信息:
	print(type(filename))
	mime.add_header('Content-Disposition', 'attachment', filename= '=?utf-8?b?' + base64.b64encode(filename) + '?=')
	# 把附件的内容读进来:
	mime.set_payload(f.read())
	# 用Base64编码:
	Encoders.encode_base64(mime)
	# 添加到MIMEMultipart:
	msg.attach(mime)

msg["From"] = _format_addr(u'HadesZ <%s>' % from_addr)
msg["To"] = _format_addr(u'小娟娟 <%s>, HadesZ <%s>' % (to_addr1, to_addr2))
msg["Subject"] = Header(u'今日景区汇总', "utf-8").encode()

server = smtplib.SMTP(smtp_server)
server.set_debuglevel(1)
server.login(from_addr, password)
server.sendmail(from_addr, [to_addr1, to_addr2], msg.as_string())
server.quit()