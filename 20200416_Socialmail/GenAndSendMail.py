# !/usr/bin/env python
# coding: utf-8
import sys

from uuid import uuid4
from random import shuffle
from pandas import read_excel
from datetime import datetime

import mylibrary.genmail as gm

from mylibrary.makezip import gen_annex
from mylibrary.makezip import make_zip
from mylibrary.insert_log import insert_log

from server.database import Database
from server.sendmail import Smtp
from server.client import Client

uuidChars = ("a", "b", "c", "d", "e", "f",
             "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s",
             "t", "u", "v", "w", "x", "y", "z", "0", "1", "2", "3", "4", "5",
             "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", "I",
             "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V",
             "W", "X", "Y", "Z")

def short_uuid():
    uuid = str(uuid4()).replace('-', '')
    result = ''
    for i in range(0,8):
        sub = uuid[i * 4: i * 4 + 4]
        x = int(sub,16)
        result += uuidChars[x % 0x3E]
    return result

def get_data(eml_file, mail_excel):
    # format eml template
    header, html = gm.get_msg(eml_file)

    # get mails from excel 
    mails = read_excel(mail_excel)
    mail_list = list(mails['Email'])
    mail_list = [x for x in mail_list if str(x) != 'nan']

    return header, html, mail_list

def insert_send_mail(eml_type, eml_file, user_group, mail_excel, sm, db, annex=[], url='http://yumail.myvnc.com'):
    # get mails from excel 
    try:
        mails = read_excel(mail_excel)
        mail_list = list(mails['Email'])
        mail_list = [x for x in mail_list if str(x) != 'nan']
    except:
        mail_list = db.get_recipients_by_group(user_group)

    header, html = gm.get_msg(eml_file)
    subject_name, mailfrom = header[0], header[2]

    # mysql: insert eml data : type_id, type, subject
    try:
        eml_type_id = str(short_uuid())
        db.insert_eml(eml_data=(eml_type_id, eml_type, subject_name))
    except:
        eml_type_id = db.get_eml_id(eml_data=(eml_type, subject_name))

    html = html.replace('my_eml_type_id', eml_type_id)
    html = html.replace('my_receive_action_url', url)
    
    # mysql: insert mail_id, mail_addr
    file_in, file_out, zip_name = annex, '測資/附件/', '測資/annex.zip'
    
    shuffle(mail_list)
    for recipient in mail_list:
        try:
            mail_id = str(short_uuid())
            db.insert_recipient(user_data=(mail_id, user_group, recipient))
        except:
            mail_id = db.get_recipient_id(user_data=(user_group, recipient))

        current_html = html.replace('my_mail_id', mail_id)
        current_html = current_html.replace('my_recipient_mail', recipient)

        # create zip annex file
        gen_annex(file_in, file_out, eml_type_id, mail_id, url)
        make_zip(file_out, zip_name)

        # create mail content
        mail_msg = gm.update_eml(eml_file, recipient, current_html, annex=[zip_name])
        
        # send mail
        sm.send(mail_msg.as_string(), mailfrom, recipient)

        # insert log to datebase
        insert_log(params=[eml_type_id, '00', mail_id], db=db)

        print(eml_type_id, subject_name)
        print(mail_id, recipient)
        print()
        

if __name__ == '__main__':

    eml_type, eml_file = '金融', '測資/test.eml'
    user_group, mail_excel = '智慧資安', '測資/recipients.xlsx'
    annex_file, date = '測資/風險金融商品比較表.doc', '2020-06-03 11:10:00'
    smtp_data = ['yumail.myvnc.com', '465', 'yucc@yumail.myvnc.com', 'yucc1234']
    db_data = ['yumail.myvnc.com', '3306', 'socialmails', 'socialmails123', 'socialmails']

    mailserver = Smtp(smtp_data[0], int(smtp_data[1]), smtp_data[2], smtp_data[3])
    database = Database(db_data[0], int(db_data[1]), db_data[2], db_data[3], db_data[4])
    insert_send_mail(eml_type, eml_file, user_group, mail_excel, mailserver, database, annex=annex_file)
    mailserver.close()
    database.__disconnect__()
    
    # client = Client()
    # client.send(smtp_data, db_data, eml_type, eml_file, user_group, mail_excel, annex_file, date)