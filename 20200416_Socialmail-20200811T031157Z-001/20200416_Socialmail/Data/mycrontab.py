import os

from crontab import CronTab

def set_cron(params, times='', every=False):
    cron = CronTab(user=True)
    command = 'python 0416_generate_and_send_mail.py '
    params = ' '.join(params)
    job = cron.new(command=command+params)
    # times = min, hr, day, mon, dow    (type: string)
    job.setall(times[0], times[1], times[2], times[3], times[4])
    cron.write()

if __name__ == '__main__':
    
    eml_file = '測資/即將到來的金融崩潰有徵兆.eml'
    mail_excel = '測資/Mails.xlsx'
    params = [eml_file, mail_excel]

    times = [None, None, '3,4,5', '11', None]
    subprocess.call('time')
    # set_cron(params, times)

