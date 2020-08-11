import smtplib

class Smtp(object):

    def __init__(self, host='127.0.0.1', port=465, user='yucc@yumail.myvnc.com', password='yucc1234'):
        try:
            self.smtp = smtplib.SMTP_SSL(host, port)
        except:
            self.smtp = smtplib.SMTP(host, port)
        
        self.smtp.login(user, password)

    def close(self):
        self.smtp.quit()

    def send(self, msg, mailfrom, rcptto):
        try:
            self.smtp.sendmail(mailfrom, rcptto, msg)
        except Exception as e:
            print('Failed to send mail.')
            # print(str(e))
            return str(e)
        else:
            print('Succeeded to send mail.')
            return None
