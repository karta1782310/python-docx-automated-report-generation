import os
import sys
import socket
import server.buffer as bf

class Client:
    def __init__(self, host='127.0.0.1',  port=6666):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            s.connect((host, port))
        except socket.error as msg:
            print(msg)
            sys.exit(1)

        self.sbuf = bf.Buffer(s)

        # Hi from server
        ehlo = self.sbuf.get_utf8()
        print(ehlo)


    def send(self, eml_type, eml_name, user_group, mail_excel, annex, datetime):
        self.sbuf.put_utf8(eml_type)
        self.sbuf.put_utf8(eml_name.split('/')[-1])
        self.sbuf.put_utf8(user_group)
        self.sbuf.put_utf8(mail_excel.split('/')[-1])
        self.sbuf.put_utf8(annex)
        self.sbuf.put_utf8(datetime)

        for file_name in [eml_name, mail_excel, annex]:
            file_size = os.path.getsize(file_name)
            self.sbuf.put_utf8(str(file_size))

            with open(file_name, 'rb') as f:
                self.sbuf.put_bytes(f.read())
            print(file_name,' Sent')


if __name__ == "__main__":
    eml_type, eml = '金融', 'test.eml'
    user_group, recipients = '智慧資安', 'recipients.xlsx'
    annex, date = '風險金融商品比較表.doc', '2020-05-25 14:07:00'

    client = Client()
    client.send(eml_type, eml, user_group, recipients, annex, date)