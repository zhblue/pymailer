#!/usr/bin/python3
# -*- coding: UTF-8 -*-

import smtplib
import poplib
import requests
import email
import urllib.parse
from email.parser import Parser
from email.utils import parseaddr
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header
from email.header import decode_header
import subprocess


def execute_cmd(cmd):
    p = subprocess.Popen(cmd, shell=True,
                         stdin=subprocess.PIPE,
                         stdout=subprocess.PIPE,
                         stderr=subprocess.PIPE)
    stdout, stderr = p.communicate()

    if p.returncode != 0:
        return stderr.decode()
    return stdout.decode()

def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos+8:].strip()
    return charset


def get_content(msg):
    for part in msg.walk():
        content_type = part.get_content_type()
        charset = guess_charset(part)
        #如果有附件，则直接跳过
        if part.get_filename()!=None:
            continue
        email_content_type = ''
        content = ''
        if content_type == 'text/plain':
            email_content_type = 'text'
        elif content_type == 'text/html':
            print('html 格式 跳过')
            continue #不要html格式的邮件
            email_content_type = 'html'
        if charset:
            try:
                content = part.get_payload(decode=True).decode(charset)
            except AttributeError:
                print('type error')
            except LookupError:
                print("unknown encoding: utf-8")
        if email_content_type =='':
            continue
            #如果内容为空，也跳过
        return content


def print_info(msg, indent=0):
    if indent == 0:
        for header in ['From', 'To', 'Subject']:
            value = msg.get(header, '')
            if value:
                if header=='Subject':
                    value = decode_str(value)
                else:
                    hdr, addr = parseaddr(value)
                    name = decode_str(hdr)
                    value = u'%s <%s>' % (name, addr)
            print('%s%s: %s' % ('  ' * indent, header, value))
    if (msg.is_multipart()):
        parts = msg.get_payload()
        for n, part in enumerate(parts):
            print('%spart %s' % ('  ' * indent, n))
            print('%s--------------------' % ('  ' * indent))
            print_info(part, indent + 1)
    else:
        content_type = msg.get_content_type()
        if content_type=='text/plain' or content_type=='text/html':
            content = msg.get_payload(decode=True)
            charset = guess_charset(msg)
            if charset:
                content = content.decode(charset)
            print('%sText: %s' % ('  ' * indent, content + '...'))
        else:
            print('%sAttachment: %s' % ('  ' * indent, content_type))


def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value
# 第三方 SMTP 服务
mail_host="smtp.qiye.aliyun.com"  #设置服务器
pop3_host="pop.qiye.aliyun.com"  #设置服务器
mail_user="you@yourdomain.com"    #用户名
mail_pass="yourpassword"   #口令
sender = mail_user

server=poplib.POP3(pop3_host)
server.user(mail_user)
server.pass_(mail_pass)

smtpObj = smtplib.SMTP()
smtpObj.connect(mail_host, 25)    # 25 为 SMTP 端口号
smtpObj.login(mail_user,mail_pass)

resp, mails, octets = server.list()
print(mails)
# 获取最新一封邮件, 注意索引号从1开始:
index = len(mails)
while index >0 :
        resp, lines, octets = server.retr(index)
        # lines存储了邮件的原始文本的每一行,
        # 可以获得整个邮件的原始文本:
        msg_content = b'\r\n'.join(lines).decode('utf-8')
        # 稍后解析出邮件:
        msg = Parser().parsestr(msg_content)
        #print_info(msg)
        cmd=decode_str(msg.get("subject"))
        if cmd.lower()=="download":
                url=get_content(msg)
        else:
                url=cmd
        if not url.startswith('http'):
                url="https://www.bing.com/search?q="+urllib.parse.quote(url)
        print("GET HTTP :"+url)
        hdr,mail_from=parseaddr(msg.get("from"))
        print("Get mail from: "+mail_from);
        # 可以根据邮件索引号直接从服务器删除邮件:
        print("delete "+str(index))
        server.dele(index)
        index=index-1
        # 关闭连接
        if cmd.lower()=="cmd" :
                tp=""
                url=get_content(msg)
                mail_msg="来信收到,具体情况如下:<br><pre>"+execute_cmd(url).replace('\n',"<br>\n")+"<pre>"
        else:
                headers = {'User-Agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3202.89 Safari/537.36'}
                cook = {"Cookie":'BAIDUID=FE0F97F1FC37C47792091A2523CD945F:FG=1; HMACCOUNT=CC6D0E280C842123'}
                lit = ['caj','htm','html','pdf', 'zip', 'rar', 'gz','exe','msi','xls','doc','xlsx','docx','txt','jpg','jpeg','png','gif','mp4','mp3','mid']
                for tp in lit:
                        if url.endswith(tp) :
                                urllib.request.urlretrieve(url, "attach."+tp)
                                break;
                        else:
                                tp="";
                res=requests.get(url,cookies=cook,headers=headers)
                res.encoding='utf-8'
                mail_msg=res.text
                print(res)

        if tp=="":
                message = MIMEText(mail_msg, 'html', 'utf-8')
        else:
                message = MIMEMultipart()
                part2=MIMEText("您的来信收到,关于所需的文件详见文本正文或附件", 'html', 'utf-8')
                part=MIMEApplication(open("attach."+tp,"rb").read())
                part.add_header('Content-Disposition', 'attachment', filename="attach."+tp)
                message.attach(part2)
                message.attach(part)
        message['From'] = mail_user
        message['To'] = mail_from

        subject = 'Python SMTP 邮件测试'
        message['Subject'] = "RE:关于"+url

        receivers = [mail_from]  # 接收邮件，可设置为你的QQ邮箱或者其他邮箱
        try:
            smtpObj.sendmail(sender, receivers, message.as_string())
            print ("邮件发送成功:"+mail_msg)
        except smtplib.SMTPException:
            print ("Error: 无法发送邮件" )

server.quit()
