import re
import email

from email.utils import formataddr, parseaddr
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.header import Header, decode_header, make_header

def parse_header(msg):
    try:
        mail_sub = str(make_header(decode_header(msg['Subject'])))
    except:
        mail_sub = msg['Subject']
    mail_to = msg['to']

    realname_from, mail_from = parseaddr(msg['from'])
    try:
        bin_name = decode_header(realname_from)
        realname_from = bin_name[0][0].decode(bin_name[0][1])
    except:
        pass

    return [mail_sub, realname_from, mail_from, mail_to]

def get_msg(FilePath):
    fp = open(FilePath, 'rb')
    msg = email.message_from_binary_file(fp)
    
    header, html = parse_header(msg), ''
    
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
    
        if part.get_content_type() == "text/html":
            html = html + part.get_payload(decode=True).decode()
            
    return header, html
    
def update_eml(eml_file, recipient, html, sender=[], annex=[]):
    with open(eml_file) as fp:
        # Create a text/plain message
        org_msg = fp.read()
    
    msg_list = org_msg.split('--===============')
    
    if sender and 'From: ' in msg_list[0]:
        new_header = msg_list[0].split('From: ')
        new_from = new_header[-1].split('\n', 1)
        msg_list[0] = 'From: '.join(new_header[:-1]) + 'From: ' + formataddr(sender) + '\n' + new_from[1]

    if 'To: ' in msg_list[0]:
        new_header = msg_list[0].split('To: ')
        new_to = new_header[-1].split('\n', 1)
        msg_list[0] = 'To: '.join(new_header[:-1]) + 'To: ' + recipient + '\n' + new_to[1]

    for idx, text in enumerate(msg_list):
        if 'text/html' in text:
            new_html = text.split('\n', 1)[0]+'\n'
            new_html = new_html + MIMEText(html, 'html', 'utf-8').as_string()
            msg_list[idx] = new_html

    new_msg = '--==============='.join(msg_list)
    msg = email.message_from_string(new_msg)

    if annex:
        msg_add_annex(msg, annex)

    return msg

def gen_test_eml(header, message):
    msg = MIMEMultipart('related')
    msg['Subject'] = Header(header[0], 'utf-8')
    msg['From'] = formataddr([header[1], header[2]], 'utf-8')
    msg['To'] = formataddr(['', header[3]], 'utf-8')
    msg['Disposition-Notification-To'] = formataddr([header[1], header[2]], 'utf-8')
    msg.attach(MIMEText(message, 'plain', 'utf-8'))

    return msg

def gen_eml(header, html, annex=[]):
    msg = MIMEMultipart('related')
    msg['Subject'] = Header(header[0], 'utf-8')
    msg['From'] = formataddr([header[1], header[2]], 'utf-8')
    msg['To'] = formataddr(['', header[3]], 'utf-8')
    
    msgAlternative = MIMEMultipart('alternative')
    msgAlternative.attach(MIMEText(html, 'html', 'utf-8'))
    msg.attach(msgAlternative)
    
    if annex:
        msg_add_annex(msg, annex)

    return msg

def msg_add_annex(msg, annex_list=[]):
    pic_idx = 0
    for filename in annex_list:
        if '.jpg' in filename or '.png' in filename:
            with open(filename, 'rb') as fp:        
                msgImage = MIMEImage(fp.read())
                msgImage.add_header('Content-ID', '<my_picture_'+str(pic_idx)+'>')
                msg.attach(msgImage)
            pic_idx = pic_idx + 1
    
        if '.zip' in filename:
            with open(filename, 'rb') as f:
                file1 = MIMEText(f.read(), 'base64', 'utf-8')
                file1["Content-Type"] = 'application/octet-stream'
                file1["Content-Disposition"] = 'attachment; filename="'+filename.split('/')[-1]+'"' 
                msg.attach(file1)

# from email import generator
# def SaveToFile(self, msg):
#     with open(self.out_file, 'w') as outfile:
#         gen = generator.Generator(outfile)
#         gen.flatten(msg)
