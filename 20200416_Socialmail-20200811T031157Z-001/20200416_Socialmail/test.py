# from pandas import DataFrame
# from server.database import Database
# from server.sendmail import Smtp
# import mylibrary.genmail as gm
 
# smtp_data = ['CAS-HT02.systex.tw', '25', '2000089', 'Cdec76461110']
# mailserver = Smtp(smtp_data[0], int(smtp_data[1]), smtp_data[2], smtp_data[3])
# mail_msg = gm.gen_test_eml(['Test Email', '寄件人', 'anthonycheng@systex.com', 'karta1782310@gmail.com'], 'This is a test mail.')
# mailserver.send(mail_msg.as_string(), 'anthonycheng@systex.com', 'karta1782310@gmail.com')
# mailserver.close()

# db = Database()
# logs = db.get_logs()
# # df = DataFrame(logs)
# # print(df)
# col_lst = list(logs[0].keys())
# print(col_lst)

test ='asdfghjk'
print(test[-1:0:-1])

