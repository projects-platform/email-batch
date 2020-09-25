#!/usr/bin/env python3
# -*- encoding: utf-8 -*-
'''
@File    :   read_email.py
@Time    :   2020/09/20 12:41:00
@Author  :   Silence H_VK
@Version :   1.0
@Contact :   hvkcoder@gmail.com
@Desc    :   None
'''

# here put the import lib

import poplib
import email
import configparser
import base64
import os
from bs4 import BeautifulSoup
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from email.message import Message


def print_info(msg, indent=0):
    if indent == 0:
        for header in ['From', 'To', 'Subject']:
            value = msg.get(header, '')
            if value:
                if header == 'Subject':
                    value = decode_str(value)
                else:
                    hdr, addr = parseaddr(value)
                    name = decode_str(hdr)
                    value = u'%s <%s>' % (name, addr)
            print('%s%s: %s' % ('  ' * indent, header, value))
    for part in msg.walk():
        if not part.is_multipart():
            content_type = part.get_content_type()
            if content_type == 'application/octet-stream':
                file_name = part.get_filename()
                if file_name:
                    file_name = decode_str(file_name)
                    if(file_name.endswith('.eml')):
                        eml_msg = email.message_from_bytes(
                            part.get_payload(decode=True))
                        payer_name = ""
                        file_name = ""
                        for eml_part in eml_msg.walk():
                            if not eml_part.is_multipart():
                                content_type = eml_part.get_content_type()
                                if(content_type == "text/html"):
                                    payer_name = get_payer_name(eml_part)
                                elif(content_type == "text/plain"):
                                    file_name = save_attachment(
                                        payer_name, eml_part)
                        print("Attachment：" + file_name)


def get_payer_name(eml_part):
    html_content = eml_part.get_payload(
        decode=True)
    soup = BeautifulSoup(
        html_content, features="html.parser")
    payer = soup.find('td', width="538").select('p')[
        7].string
    return payer[payer.find('：')+1:]


def save_attachment(payer_name, eml_part):
    file_name = eml_part.get_filename()
    file_name = "%s-%s" % (payer_name, file_name)
    with open("%s/%s" % (base_save_path, file_name), 'wb') as file:
        file.write(eml_part.get_payload(decode=True))
    return file_name


def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value


def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset


if __name__ == "__main__":
    # 从配置文件中读取邮箱配置
    try:
        Config = configparser.ConfigParser()
        Config.read("config.ini")
        pop_server = Config.get("email", "pop_server")
        email_address = Config.get("email", "email_address")
        email_password = Config.get("email", "email_password")
        base_save_path = Config.get("email", "base_save_path")
        if(base_save_path == None or base_save_path == ""):
            print("❌ 需要配置附件存储地址")
            exit(1)
        if not os.path.isdir(base_save_path):
            os.makedirs(base_save_path)

    except Exception as e:
        print("❌ 读取配置文件失败，请查看配置")
        exit(1)

    # 连接服务器
    try:
        pop3Server = poplib.POP3_SSL(pop_server)
    except OSError as e:
        print('❌ 连接服务器失败，请检查服务器地址或网络连接。')
        exit(1)

    # 验证用户名
    try:
        pop3Server.user(email_address)
    except Exception as e:
        print('❌ 邮箱地址不存在')
        exit(1)

    # 验证密码
    try:
        pop3Server.pass_(email_password)
    except Exception as e:
        print('❌ 邮箱密码正确')
        exit(1)

    # 输出欢迎符号
    print(f'🎉 {pop3Server.getwelcome().decode()}')

    # 邮件数量和总大小
    mail_quantity, mail_total_size = pop3Server.stat()
    print(f'👉 邮件总数:{mail_quantity}')
    print(f'👉 邮件总大小:{mail_total_size}')

    # 倒序读取邮件内容
    response, mail_list, octets = pop3Server.list()
    email_count = len(mail_list)
    for mail_info in reversed(mail_list):
        mail_number = int(bytes.decode(mail_info).split()
                          [0])  # 从mail_list解析邮件编号
        print(
            f'------------------------------  正在获取邮件({(email_count - mail_number)+ 1}/{email_count})  ------------------------------')
        resp, lines, octets = pop3Server.retr(mail_number)
        msg_content = b'\r\n'.join(lines).decode('utf-8')
        msg = Parser().parsestr(msg_content)
        print_info(msg)

    print("✅ 读取邮件完成")
    pop3Server.close()
