# coding: utf-8
import os
import sys
import socket
import threading

import buffer

from time import sleep
from scheduler import scheduler

class Receiver :

    def __init__(self, host='0.0.0.0', port=6666):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            s.bind((host, port))
            s.listen(5)
        except socket.error as msg:
            print(msg)
            sys.exit(1)
        
        while True:
            conn, addr = s.accept()
            connbuf = buffer.Buffer(conn)
            recv_thread = threading.Thread(target = self.deal_data, args=(connbuf, addr))
            recv_thread.start()

    def deal_data(self, connbuf, addr):
        print()
        print("Got a connection from ", addr)
        absolute_path = '/var/www/socialmails/schedule_server/'

        connbuf.put_utf8('Hi, Welcome to the server!')
        smtp_data = connbuf.get_utf8()
        db_data = connbuf.get_utf8()
        eml_type = connbuf.get_utf8()
        eml_name = absolute_path+'eml/'+connbuf.get_utf8().split('/')[-1]
        user_group = connbuf.get_utf8()
        mail_excel = absolute_path+'excel/'+connbuf.get_utf8().split('/')[-1]
        annex = absolute_path+'annex/'+connbuf.get_utf8().split('/')[-1]
        url = connbuf.get_utf8()
        datetime = connbuf.get_utf8()
        
        absolute_path = '/var/www/socialmails/schedule_server/'

        for file_name in [eml_name, mail_excel, annex]:
            file_size = int(connbuf.get_utf8())
            print('file size: ', file_size )

            with open(file_name, 'wb') as f:
                remaining = file_size
                while remaining:
                    chunk_size = 4096 if remaining >= 4096 else remaining
                    chunk = connbuf.get_bytes(chunk_size)
                    if not chunk: break
                    f.write(chunk)
                    remaining -= len(chunk)
                if remaining:
                    print(file_name,' incomplete.  Missing',remaining,'bytes.')
                else:
                    print(file_name,' received successfully.')

        print('All data ({0}, {1}, {2})'.format(smtp_data, db_data, url))
        print()
        scheduler(datetime, [smtp_data, db_data, eml_type, eml_name, user_group, mail_excel, annex, url])
        
        
if __name__ == "__main__":
    receiver = Receiver()
