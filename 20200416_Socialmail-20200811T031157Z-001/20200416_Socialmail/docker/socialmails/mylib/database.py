import pymysql

class Database():

    def __init__(self, host='dev_database', port=3306, user='socialmails',
                       password='socialmails123', db='socialmails', charset='utf8'):
        self.host = host
        self.port = port
        self.user = user
        self.password = password
        self.db = db
        self.charset = charset

    def __connect__(self):
        self.con = pymysql.connect(host=self.host, port=self.port, user=self.user,
                                   password=self.password, db=self.db, charset=self.charset)
        self.cur = self.con.cursor(cursor = pymysql.cursors.DictCursor)

    def __disconnect__(self):
        self.con.commit()
        self.cur.close()
        self.con.close()

    def fetch(self, sql):
        self.__connect__()
        self.cur.execute(sql)
        result = self.cur.fetchall()
        self.__disconnect__()
        return result

    def execute(self, sql):
        self.__connect__()
        self.cur.execute(sql)
        self.__disconnect__()

    def get_recipients_by_group(self, user_group):
        self.__connect__()
        sql = "select user_email from mail_recipients where user_group=%s"
        self.cur.execute(sql, (user_group))
        result = self.cur.fetchall()
        mail_list = [mail['user_email'] for mail in result]
        self.__disconnect__()
        return mail_list

    def insert_eml(self, eml_data=()):
        self.__connect__()
        sql = "insert into eml_files(id, type, subject) values(%s, %s, %s)"
        self.cur.execute(sql, eml_data)
        self.__disconnect__()
    
    def get_eml_id(self, eml_data=()):
        self.__connect__()
        sql = "select id from eml_files where type=%s and subject=%s"
        self.cur.execute(sql, eml_data)
        result = self.cur.fetchall()
        eml_type_id = result[0]['id']
        self.__disconnect__()
        return eml_type_id

    def insert_recipient(self, user_data=()):
        self.__connect__()
        sql = "insert into mail_recipients(id, user_group, user_email) values(%s, %s, %s)"
        self.cur.execute(sql, user_data)
        self.__disconnect__()

    def get_recipient_id(self, user_data=()):
        self.__connect__()
        sql = "select id from mail_recipients where user_group=%s and user_email=%s"
        self.cur.execute(sql, user_data)
        result = self.cur.fetchall()
        mail_id = result[0]['id']
        self.__disconnect__()
        return mail_id

    def insert_log(self, user_data=()):
        self.__connect__()
        self.cur.execute('alter table user_logs auto_increment = 1')
        sql = "insert into user_logs(eml_id, recipient_id, action) values (%s, %s, %s)"
        self.cur.execute(sql, user_data)
        self.__disconnect__()