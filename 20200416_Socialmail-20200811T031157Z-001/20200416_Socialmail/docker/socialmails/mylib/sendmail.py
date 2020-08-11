import smtplib

class Smtp(object):

    def __init__(self, host='dev_mailserver', port=465, user='yucc@yumail.myvnc.com', password='yucc1234'):
        self.smtp = smtplib.SMTP_SSL(host, port)
        self.smtp.login(user, password)

    def close(self):
        self.smtp.quit()

    def send(self, msg, mailfrom, rcptto):
        # msg = None
        # with open(infile) as f:
        #     msg = f.read()

        try:
            self.smtp.sendmail(mailfrom, rcptto, msg)
        except Exception as e:
            print('Failed to send mail.')
            print(str(e))
        else:
            print('Succeeded to send mail.')





