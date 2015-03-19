#!/usr/bin/python
# -*- coding: utf-8 -*-

import sys
import xlrd
import MySQLdb
import smtplib
import re

product='DRP'

def db_sele(a):
	SQL="SELECT path FROM zentao.zt_module WHERE id=%s"
	cursor.execute(SQL,[a])
	b=cursor.fetchone()
	#print b[0]
	c=b[0].split(',')
	f=''
	for d in c:
		if d!='':
			SQL="SELECT name FROM zentao.zt_module WHERE id=%s"
			cursor.execute(SQL,[d])
			e=cursor.fetchone()
			if f=='':
				f=e[0]
			else:
				f+='-'+e[0]
	return f

conn=MySQLdb.connect(host='192.168.200.18', user='root', passwd='123456', db='zentao')
cursor=conn.cursor()

SQL="set names 'GBK'"
cursor.execute(SQL)

SQL='SELECT a.id,a.module,a.title, b.spec \
FROM zentao.zt_story a, zentao.zt_storyspec b, zentao.zt_product c \
WHERE c.name=%s and a.product=c.id and a.id=b.story'
cursor.execute(SQL,[product])

#reload(sys)
#sys.setdefaultencoding('utf-8')
import xlwt
#xlwt.Book.encoding='gbk'
#w=xlwt.Workbook(encoding = 'ascii')
#w=xlwt.Workbook()
w=xlwt.Workbook(encoding='gbk')
ws=w.add_sheet('Requirment')
ws.write(0,0,'Requirment ID')
ws.write(0,1,'Requirment Name')
ws.write(0,2,'Module ID')
ws.write(0,3,'Requirment Description')
ws.write(0,4,'Parent Module')

i=0
for id,module,title,desc in cursor.fetchall():
	i+=1
	#print id,module
	desc = re.sub('<[^>]*?>','',desc)
	#desc.strip()

	ws.write(i,0,id)
	ws.write(i,1,title)
	ws.write(i,2,module)

	#a=''
	#for b in desc.split():
	#	a+=b
	#	a+='\n'
	#desc=a
	#style=xlwt.easyxf('align: wrap on')
	#ws.write(i,3,desc,style)
	ws.write(i,3,desc)

	if module!= 0 :
		module_list=db_sele(module)
		#print id,title
		#print module,desc
		#print module_list

		ws.write(i,4,module_list)
		#break
	else:
		ws.write(i,4,'/')
	#break

w.save('requirment.xls')
cursor.close()
conn.close()


import email.MIMEMultipart
import email.MIMEText
import email.MIMEBase
import os.path

From="codereview_cccis@163.com"
#To="yfeng@cccis.com"
To="hyi@cccis.com"
file_name = "./requirment.xls"

server = smtplib.SMTP("smtp.163.com")
server.login("codereview_cccis@163.com","abcd@1234") #仅smtp服务器需要验证时

# 构造MIMEMultipart对象做为根容器
main_msg = email.MIMEMultipart.MIMEMultipart()

# 构造MIMEText对象做为邮件显示内容并附加到根容器
text_msg = email.MIMEText.MIMEText("Requirment list from Zentao")
main_msg.attach(text_msg)

# 构造MIMEBase对象做为文件附件内容并附加到根容器
contype = 'application/octet-stream'
maintype, subtype = contype.split('/', 1)

## 读入文件内容并格式化
data = open(file_name, 'rb')
file_msg = email.MIMEBase.MIMEBase(maintype, subtype)
file_msg.set_payload(data.read( ))
data.close( )
email.Encoders.encode_base64(file_msg)

## 设置附件头
basename = os.path.basename(file_name)
file_msg.add_header('Content-Disposition',
 'attachment', filename = basename)
main_msg.attach(file_msg)

# 设置根容器属性
main_msg['From'] = From
main_msg['To'] = To
main_msg['Subject'] = "Requirment"
main_msg['Date'] = email.Utils.formatdate( )

# 得到格式化后的完整文本
fullText = main_msg.as_string( )

# 用smtp发送邮件
try:
    server.sendmail(From, To, fullText)
finally:
    server.quit()

os.remove("./requirment.xls")

#exit()
exit()
