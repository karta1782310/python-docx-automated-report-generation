import pymysql

class Database():

    def __init__(self, host='127.0.0.1', port=3306, user='socialmails',
                       password='socialmails123', db='socialmails', charset='utf8'):
        self.host = host
        self.port = port
        self.user = user
        self.password = password
        self.db = db
        self.charset = charset
        self.__connect__()

    def __connect__(self):
        self.con = pymysql.connect(host=self.host, port=self.port, user=self.user,
                                   password=self.password, db=self.db, charset=self.charset)
        self.cur = self.con.cursor(cursor = pymysql.cursors.DictCursor)

    def __disconnect__(self):
        self.con.commit()
        self.cur.close()
        self.con.close()

    def fetch(self, sql, params=None):
        # self.__connect__()
        self.cur.execute(sql, params or ())
        result = self.cur.fetchall()
        # self.__disconnect__()
        return result

    def execute(self, sql, params=None):
        # self.__connect__()
        self.cur.execute(sql, params or ())
        # self.__disconnect__()

    def get_recipients_by_group(self, user_group):
        sql = "select user_email from mail_recipients where user_group=%s"
        result = self.fetch(sql, (user_group))
        mail_list = [mail['user_email'] for mail in result]
        return mail_list

    def insert_eml(self, eml_data=()):
        sql = "insert into eml_files(id, type, subject) values(%s, %s, %s)"
        self.execute(sql, eml_data)
    
    def get_eml_id(self, eml_data=()):
        sql = "select id from eml_files where type=%s and subject=%s"
        result = self.fetch(sql, eml_data)
        eml_type_id = result[0]['id']
        return eml_type_id

    def get_eml_data_by_id(self, eml_id=''):
        sql = "select * from eml_files where id=%s"
        result = self.fetch(sql, (eml_id))
        return result

    def insert_recipient(self, user_data=()):
        sql = "insert into mail_recipients(id, user_group, user_email) values(%s, %s, %s)"
        self.execute(sql, user_data)

    def get_recipient_id(self, user_data=()):
        sql = "select id from mail_recipients where user_group=%s and user_email=%s"
        result = self.fetch(sql, user_data)
        mail_id = result[0]['id']
        return mail_id

    def get_recipient_data_by_id(self, recipient_id=''):
        sql = "select * from mail_recipients where id=%s"        
        result = self.fetch(sql, (recipient_id))
        return result

    def insert_log(self, user_data=()):
        self.execute('alter table user_logs auto_increment = 1')
        sql = "insert into user_logs(eml_id, recipient_id, action) values (%s, %s, %s)"
        self.execute(sql, user_data)

    def get_logs(self):
        sql = "SELECT log.id, eml.type, eml.subject, usr.user_group, usr.user_email, log.date, log.action \
            FROM eml_files as eml, mail_recipients as usr, user_logs as log \
            WHERE log.eml_id=eml.id and log.recipient_id=usr.id \
            ORDER BY log.id;"
        result = self.fetch(sql)
        return result