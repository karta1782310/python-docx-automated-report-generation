# coding: utf-8
import os
import sys
import subprocess

from datetime import datetime

def format_date(time):
    # format time to 'at' needed
    dateString = time
    dateFormatter = '%Y-%m-%d %H:%M:%S'
    times = datetime.strptime(dateString, dateFormatter)
    times = times.strftime('%H:%M %d.%m.%Y')
    return times

def scheduler(date_in, args=[]):
    time = format_date(date_in)

    # get exactly path of file
    filename = '/var/www/socialmails/GenAndSendMail.py'
    date_com = 'at ' + time
    python = 'python3 '
    exec_com = python+filename+' '+' '.join(args)

    # run shell command
    p = subprocess.Popen(date_com, shell=True, stdin=subprocess.PIPE)
    p.stdin.write(exec_com.encode('utf-8'))

    # eml_file=args[2], recipients_file=args[4], annex_file=args[5]
    tmp = '\nrm '+ args[3] + '\nrm ' + args[5] + '\nrm ' + args[6]
    p.stdin.write(tmp.encode('utf-8'))
    p.stdin.close()

if __name__ == '__main__':
    dateString = '2020-04-23 10:21:00'
    
    # if sys.argv[1:]:
    #     eml_type, eml_name = sys.argv[1], sys.argv[2]
    #     user_group, mail_excel = sys.argv[3], sys.argv[4]

    scheduler(dateString, sys.argv[1:])
