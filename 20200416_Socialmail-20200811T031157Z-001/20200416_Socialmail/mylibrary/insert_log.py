import sys  
from server.database import Database

def insert_log(params, db=None):
    if not db:
        db = Database()

    eml_id = params[0]
    act = params[1]
    recipient_id = params[2]

    if act == '00':
        act = '送信'
    elif act == '01':
        act = '開信'
    elif act == '02':
        act = '點擊連結'
    elif act == '03':
        act = '開啟附件'

    db.insert_log((eml_id, recipient_id, act))

if __name__ == '__main__':
    params = sys.argv[1:]
    insert_log(params)

