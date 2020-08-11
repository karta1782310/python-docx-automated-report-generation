#!/bin/bash
# -*- coding: UTF-8 -*-
# 基本控件都在这里面
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QGridLayout, QMessageBox, QFileDialog,
                            QLabel, QLineEdit, QPushButton, QComboBox, QCheckBox, QDateTimeEdit, 
                            QTextEdit, QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView)

from PyQt5.QtGui import QPalette, QColor, QBrush
from PyQt5.QtCore import Qt, QDateTime

from pyqtgraph import GraphicsLayoutWidget, setConfigOption, setConfigOptions
import qdarkstyle, sys

import mylibrary.genmail as gm

from GenAndSendMail import insert_send_mail
from server.database import Database
from server.sendmail import Smtp
from server.client import Client
from email import generator
from pandas import DataFrame
from copy import deepcopy

class SubWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.resize(400,100)
        self.main_layout = QGridLayout()
        self.setLayout(self.main_layout)

        self.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())

        self.main_layout.addWidget(QLabel('收件人'), 0, 0, 1, 1)
        self.in_recipient = QLineEdit()
        self.main_layout.addWidget(self.in_recipient, 0, 1, 1, 5)
        self.btn_send = QPushButton('寄送')
        self.main_layout.addWidget(self.btn_send, 1, 5, 1, 1)

class MailserverUi(QMainWindow):
    def __init__(self):
        super().__init__()

        setConfigOption('background', '#19232D')
        setConfigOption('foreground', 'd')
        setConfigOptions(antialias = True)
        
        # self.resize(720,500)
        self.init_ui()
        self.data_smtp = []
        self.data_db = []
        self.data_logs = []
        self.data_temp_logs = []

        # self.sub_win = SubWindow()

        # 默認狀態欄
        self.status = self.statusBar()
        self.status.showMessage("開發者: 鄭鈺城, 聯絡資訊: anthonycheng@systex.com")
        
        # 標題欄
        self.setWindowTitle("社交郵件工程")
        self.setWindowOpacity(1) # 窗口透明度
        self.main_layout.setSpacing(0)
        
        self.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
        self.main_widget.setStyleSheet(
            """
            QComboBox::item:checked {
                height: 12px;
                border: 1px solid #32414B;
                margin-top: 0px;
                margin-bottom: 0px;
                padding: 4px;
                padding-left: 0px;
            }
            """
        )

    def init_ui(self):
        # 創建視窗主部件
        self.main_widget = QWidget()  
        # 創建主部件的網格佈局
        self.main_layout = QGridLayout()  
        # 設置窗口主部件佈局為網格佈局
        self.main_widget.setLayout(self.main_layout)  

        # 創建左側部件
        self.left_widget = QWidget()  
        self.left_widget.setObjectName('left_widget')
        self.left_layout = QGridLayout()
        self.left_widget.setLayout(self.left_layout) 

        # 創建右側部件
        self.right_widget = QWidget() 
        self.right_widget.setObjectName('right_widget')
        self.right_layout = QGridLayout()
        self.right_widget.setLayout(self.right_layout) 

        # 左側部件在第0行第0列，佔12行3列
        self.main_layout.addWidget(self.left_widget, 0, 0, 12, 3) 
        # 右側部件在第0行第3列，佔12行8列
        self.main_layout.addWidget(self.right_widget, 0, 3, 12, 8)
        # 設置視窗主部件
        self.setCentralWidget(self.main_widget) 

        # 主要功能按鈕
        self.btn_sendmail = QPushButton("發送信件")
        self.btn_sendmail.clicked.connect(self.display_send_mail)
        self.btn_smtp = QPushButton("系統設定")
        self.btn_smtp.clicked.connect(self.display_smtp_setting)
        self.btn_db = QPushButton("資料庫設定")
        self.btn_db.clicked.connect(self.display_db_setting)
        self.btn_update_eml = QPushButton("修改樣板")
        self.btn_update_eml.clicked.connect(self.display_update_eml)
        self.btn_get_logs = QPushButton("觸發明細")
        self.btn_get_logs.clicked.connect(self.display_logs)
        self.btn_download_logs = QPushButton("下載觸發明細")
        self.btn_download_logs.clicked.connect(self.logs_download)
        self.quit_btn = QPushButton("退出")
        self.quit_btn.clicked.connect(self.quit_act)

        self.left_layout.addWidget(self.btn_sendmail, 2, 0, 1, 3)
        self.left_layout.addWidget(self.btn_smtp, 3, 0, 1, 3)
        self.left_layout.addWidget(self.btn_db, 4, 0, 1, 3)
        self.left_layout.addWidget(self.btn_update_eml, 5, 0, 1, 3)
        self.left_layout.addWidget(self.btn_get_logs, 6, 0, 1, 3)
        self.left_layout.addWidget(self.btn_download_logs, 7, 0, 1, 3)
        self.left_layout.addWidget(self.quit_btn, 8, 0, 1, 3)

        # 主要功能查詢
        self.in_data = QLineEdit()
        self.in_data.setPlaceholderText("暫無")
        self.left_layout.addWidget(self.in_data, 1, 0, 1, 3)

        # 主要功能 log
        self.query_result = QTableWidget()
        self.left_layout.addWidget(self.query_result, 9, 0, 2, 3)
        self.query_result.verticalHeader().setVisible(False)

        self.right_display = GraphicsLayoutWidget()
        self.right_layout.addWidget(self.right_display, 0, 3, 12, 8)

        # 右側物件: sendmail
        self.in_eml_type = QLineEdit()
        self.in_eml_template = QLineEdit()
        self.btn_eml_browse = QPushButton('瀏覽')
        self.btn_eml_browse.clicked.connect(lambda: self.open_eml(self.in_eml_template))
        self.in_recipient_group = QLineEdit()
        self.in_recipient_excel = QLineEdit()
        self.btn_recipient_browse = QPushButton('瀏覽')
        self.btn_recipient_browse.clicked.connect(lambda: self.open_excel(self.in_recipient_excel))
        self.in_annex_file = QLineEdit()
        self.btn_annex_file = QPushButton('瀏覽')
        self.btn_annex_file.clicked.connect(lambda: self.open_word(self.in_annex_file))
        self.in_scheduler = QDateTimeEdit(QDateTime.currentDateTime())
        self.in_scheduler.setCalendarPopup(True)
        self.in_scheduler.setDisplayFormat('yyyy-MM-dd hh:mm')
        self.cb_scheduler = QCheckBox('使用')
        self.btn_sendmail_start = QPushButton('執行')
        self.btn_sendmail_start.clicked.connect(self.send_mail)

        # 右側物件: smtp
        self.in_smtp_host = QLineEdit()
        self.in_smtp_port = QLineEdit()
        self.in_smtp_user = QLineEdit()
        self.in_smtp_password = QLineEdit()
        self.cb_smtp_ssl = QCheckBox('使用')
        self.in_smtp_test = QLineEdit()
        self.btn_smtp_save = QPushButton('儲存')
        self.btn_smtp_save.clicked.connect(lambda: self.save_data(self.data_smtp))
        self.btn_smtp_test = QPushButton('測試')
        self.btn_smtp_test.clicked.connect(self.show_sub_win)

        # 右側物件: db
        self.in_db_host = QLineEdit()
        self.in_db_port = QLineEdit()
        self.in_db_user = QLineEdit()
        self.in_db_password = QLineEdit()
        self.in_db_database = QLineEdit()
        self.in_db_domain = QLineEdit()
        self.in_db_domain.setPlaceholderText('回收風險資訊動作的網址')
        self.btn_db_save = QPushButton('儲存')
        self.btn_db_save.clicked.connect(lambda: self.save_data(self.data_db))

        # 右側物件: update eml
        self.in_edit_sender = QLineEdit()
        self.in_edit_sender_name = QLineEdit()
        self.cb_edit_annex = QCheckBox('是')
        self.in_edit_annex = QLineEdit()
        self.btn_edit_annex = QPushButton('瀏覽')
        self.btn_edit_annex.clicked.connect(lambda: self.open_annex(self.in_edit_annex))
        self.in_edit_subject = QLineEdit()

        self.mail_tab = QTabWidget()
        self.mail_tab.setDocumentMode(True)
        self.mail_tab.currentChanged.connect(self.print_html)
        self.mail_tab_1 = QWidget()
        self.mail_tab_2 = QWidget()
        self.mail_tab.addTab(self.mail_tab_1, 'Html')
        self.mail_tab.addTab(self.mail_tab_2, 'Web')

        self.tab_1 = QGridLayout()        
        self.tab_2 = QGridLayout()
        self.tab_1.setContentsMargins(0,0,0,0)
        self.tab_2.setContentsMargins(0,0,0,0)
        self.mail_tab_1.setLayout(self.tab_1) 
        self.mail_tab_2.setLayout(self.tab_2)
        self.in_edit_html = QTextEdit()
        self.in_edit_web = QWebEngineView()
        self.tab_1.addWidget(self.in_edit_html, 1, 1, 1, 1)
        self.tab_2.addWidget(self.in_edit_web, 1, 1, 1, 1)

        self.btn_edit_eml_reset = QPushButton('清除')
        self.btn_edit_eml_reset.clicked.connect(self.eml_reset)
        self.btn_edit_eml_read = QPushButton('讀取')
        self.btn_edit_eml_read.clicked.connect(self.eml_open)
        self.btn_edit_eml_save = QPushButton('儲存')
        self.btn_edit_eml_save.clicked.connect(self.eml_save)

        # 右側物件: logs
        self.tbw_logs = QTableWidget()
        self.tbw_logs.verticalHeader().setVisible(False)
        self.cmb_logs_choice = QComboBox()
        self.in_logs_data = QLineEdit()
        self.in_logs_data.setPlaceholderText("輸入資料")
        self.btn_logs_search = QPushButton('執行')
        self.btn_logs_search.clicked.connect(self.logs_change)


    def display_send_mail(self):
        self.clear_layout(self.right_layout)

        labels = [ "信件類型 :", "信件模板 :", "    收件人群組 :", "收件人資料 :", '附件資料 :',"設定排程 :"]
        for i, label in enumerate(labels):
            self.right_layout.addWidget(QLabel(label), i, 3, 1, 1, Qt.AlignRight)

        self.right_layout.addWidget(self.in_eml_type, 0, 4, 1, 7)
        self.right_layout.addWidget(self.in_eml_template, 1, 4, 1, 6)
        self.right_layout.addWidget(self.btn_eml_browse, 1, 10, 1, 1)
        self.right_layout.addWidget(self.in_recipient_group, 2, 4, 1, 7)
        self.right_layout.addWidget(self.in_recipient_excel, 3, 4, 1, 6)
        self.right_layout.addWidget(self.btn_recipient_browse, 3, 10, 1, 1)
        self.right_layout.addWidget(self.in_annex_file , 4, 4, 1, 6)
        self.right_layout.addWidget(self.btn_annex_file, 4, 10, 1, 1)
        self.right_layout.addWidget(self.in_scheduler, 5, 4, 1, 6)
        self.right_layout.addWidget(self.cb_scheduler, 5, 10, 1, 1)
        self.right_layout.addWidget(self.btn_sendmail_start, 6, 9, 1, 2)

    def display_smtp_setting(self):
        self.clear_layout(self.right_layout)
        
        # 在右邊新增物件 
        labels = ["SMTP HOST :", "SMTP PORT :", "SMTP 帳號 :", "SMTP 密碼 :", "SMTP SSL :", "  測試信件內容 :"]
        for i, label in enumerate(labels):
            self.right_layout.addWidget(QLabel(label), i, 3, 1, 1, Qt.AlignRight)

        self.right_layout.addWidget(self.in_smtp_host, 0, 4, 1, 7)
        self.right_layout.addWidget(self.in_smtp_port, 1, 4, 1, 7)
        self.right_layout.addWidget(self.in_smtp_user, 2, 4, 1, 7)
        self.right_layout.addWidget(self.in_smtp_password, 3, 4, 1, 7)
        self.right_layout.addWidget(self.cb_smtp_ssl, 4, 4, 1, 7)
        self.right_layout.addWidget(self.in_smtp_test, 5, 4, 1, 7)
        self.right_layout.addWidget(self.btn_smtp_save, 6, 9, 1, 2)
        self.right_layout.addWidget(self.btn_smtp_test, 6, 7, 1, 2)
    
    def display_db_setting(self):
        self.clear_layout(self.right_layout)
        
        # 在右邊新增物件 
        labels = ["資料庫 HOST :", "資料庫 PORT :", "資料庫 帳號 :", "資料庫 密碼 :", "使用資料庫名稱 :", "回收網址 :"]
        for i, label in enumerate(labels):
            self.right_layout.addWidget(QLabel(label), i, 3, 1, 1, Qt.AlignRight)

        self.right_layout.addWidget(self.in_db_host, 0, 4, 1, 7)
        self.right_layout.addWidget(self.in_db_port, 1, 4, 1, 7)
        self.right_layout.addWidget(self.in_db_user, 2, 4, 1, 7)
        self.right_layout.addWidget(self.in_db_password, 3, 4, 1, 7)
        self.right_layout.addWidget(self.in_db_database, 4, 4, 1, 7)
        self.right_layout.addWidget(self.in_db_domain, 5, 4, 1, 7)
        self.right_layout.addWidget(self.btn_db_save, 6, 9, 1, 2)  

    def display_update_eml(self):
        self.clear_layout(self.right_layout)

        labels = ["寄件人 :", "寄件人名稱 :", "  是否加入附件 :", "附件名稱 :", "主旨 :", "內容 :"]
        for i, label in enumerate(labels):
            self.label = QLabel(label)
            self.right_layout.addWidget(self.label, i, 3, 1, 1, Qt.AlignRight)
        
        self.right_layout.addWidget(self.in_edit_sender, 0, 4, 1, 7)
        self.right_layout.addWidget(self.in_edit_sender_name, 1, 4, 1, 7)
        self.right_layout.addWidget(self.cb_edit_annex, 2, 4, 1, 7)
        self.right_layout.addWidget(self.in_edit_annex, 3, 4, 1, 6)
        self.right_layout.addWidget(self.btn_edit_annex, 3, 10, 1, 1)
        self.right_layout.addWidget(self.in_edit_subject, 4, 4, 1, 7)
        self.right_layout.addWidget(self.mail_tab, 5, 4, 6, 7)
        self.right_layout.addWidget(self.btn_edit_eml_reset, 11, 5, 1, 2)
        self.right_layout.addWidget(self.btn_edit_eml_read, 11, 7, 1, 2)
        self.right_layout.addWidget(self.btn_edit_eml_save, 11, 9, 1, 2)

    def display_logs(self):
        self.data_temp_logs = []
        self.tbw_logs.setRowCount(0)
        self.clear_layout(self.right_layout)
        self.right_layout.addWidget(self.tbw_logs, 1, 3, 11, 8)
        self.right_layout.addWidget(QLabel('查詢 :'), 0, 3, 1, 1)
        self.right_layout.addWidget(self.cmb_logs_choice, 0, 4, 1, 2)
        self.right_layout.addWidget(self.in_logs_data, 0, 6, 1, 3)
        self.right_layout.addWidget(self.btn_logs_search, 0, 9, 1, 2)

        try:
            db = Database(self.data_db[0], int(self.data_db[1]), self.data_db[2], self.data_db[3], self.data_db[4]) if self.data_db[:5] else Database()
            self.data_logs = db.get_logs()
            self.data_temp_logs =  deepcopy(self.data_logs)
            
            if self.data_logs:
                row_num = len(self.data_logs)
                col_num = len(self.data_logs[0])
                col_lst = list(self.data_logs[0].keys())
                self.cmb_logs_choice.clear()
                self.cmb_logs_choice.addItems(col_lst)

                self.tbw_logs.setRowCount(row_num)  
                self.tbw_logs.setColumnCount(col_num)
                self.tbw_logs.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
                self.tbw_logs.setHorizontalHeaderLabels(col_lst)

                for i in range(row_num):
                    row_data = list(self.data_logs[i].values())
                    for j in range(col_num):
                        temp_data = row_data[j]
                        item = QTableWidgetItem(str(temp_data))
                        item.setForeground(QBrush(QColor(144, 182, 240)))
                        self.tbw_logs.setItem(i, j, item)
        except:
            QMessageBox.warning(self, 'Failed!', '資料庫連結失敗！', QMessageBox.Ok)
        else:
            db.__disconnect__()

    def get_items_from_layout(self, layout):
        return [layout.itemAt(i).widget() for i in range(layout.count())]

    def save_data(self, data):
        items = self.get_items_from_layout(self.right_layout)
        data.clear()

        try:
            for item in items:
                if type(item) == type(QLineEdit()):
                    data.append(item.text())
                elif type(item) == type(QCheckBox()):
                    data.append(item.isChecked())      

            QMessageBox.information(self, 'Success!', '儲存成功！', QMessageBox.Ok)  
        except:
            QMessageBox.warning(self, 'Failed!', '儲存失敗！', QMessageBox.Ok)

        print(data)

    def clear_layout(self, layout):
        for i in reversed(range(layout.count())): 
            layout.itemAt(i).widget().setParent(None)

    def open_eml(self, obj):
        file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", "Eml Files (*.eml)")
        obj.setText(file_name)

    def open_excel(self, obj):
        file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", "Excel Files (*.xlsx)")
        obj.setText(file_name)

    def open_word(self, obj):
        file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", "Word Files (*.doc *.docx)")
        obj.setText(file_name)

    def open_annex(self, obj):
        file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", "Annex Files (*.jpg *.png *.zip)")
        org_files = obj.text()
        all_files = org_files + ',' + file_name if org_files else file_name
        obj.setText(all_files)

    def print_html(self, index):
        if index:
            self.in_edit_web.setHtml(self.in_edit_html.toPlainText())

    def send_mail(self):
        eml_type = self.in_eml_type.text()
        eml_file = self.in_eml_template.text()
        user_group = self.in_recipient_group.text()
        mail_excel = self.in_recipient_excel.text()
        annex_file = self.in_annex_file.text()
        url = self.data_db[5] if self.data_db else 'http://yumail.myvnc.com'

        try:
            if self.cb_scheduler.isChecked():
                my_time = self.in_scheduler.text()+':00'

                client = Client()
                client.send(self.data_smtp[:4], self.data_db[:5], eml_type, eml_file, user_group, mail_excel, annex_file, url, my_time)
                
                QMessageBox.information(self, 'Success!', '排程設定成功!', QMessageBox.Ok)
            else:
                sm = Smtp(self.data_smtp[0], int(self.data_smtp[1]), self.data_smtp[2], self.data_smtp[3]) if self.data_smtp else Smtp()
                db = Database(self.data_db[0], int(self.data_db[1]), self.data_db[2], self.data_db[3], self.data_db[4]) if self.data_db else Database()

                insert_send_mail(eml_type, eml_file, user_group, mail_excel, sm, db, annex=annex_file, url=url)
        
                sm.close()
                db.__disconnect__()
        
                QMessageBox.information(self, 'Success!', '信件寄出成功!', QMessageBox.Ok)
        except:
            QMessageBox.warning(self, 'Failed!', '信件寄出失敗!', QMessageBox.Ok)

    def show_sub_win(self):
        if self.data_smtp:
            self.sub_win = SubWindow()
            self.sub_win.btn_send.clicked.connect(self.send_test)
            self.sub_win.show()
        else:
            QMessageBox.warning(self, 'Failed!', '請確認有無 SMTP 資料!', QMessageBox.Ok)
            
    def send_test(self):
        try:
            if self.data_smtp:
                mailserver = Smtp(self.data_smtp[0], int(self.data_smtp[1]), self.data_smtp[2], self.data_smtp[3])
                mail_msg = gm.gen_test_eml(['Test Email', '測試寄件人', self.data_smtp[2], self.sub_win.in_recipient.text()], self.data_smtp[5])
                error = mailserver.send(mail_msg.as_string(), self.data_smtp[2], self.sub_win.in_recipient.text())
                mailserver.close()
                if error:
                    QMessageBox.warning(self, 'Warning!', '信件寄出成功!\nWaning: '+error, QMessageBox.Ok)
                else:
                    QMessageBox.information(self, 'Success!', '信件寄出成功!', QMessageBox.Ok)
                self.sub_win.in_recipient.clear()
        except:
            QMessageBox.warning(self, 'Failed!', '信件寄出失敗!', QMessageBox.Ok)
            
    def eml_open(self):
        self.in_edit_html.clear()
        file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", "Eml Files (*.eml)")
        if not file_name:
            return
            
        header, html = gm.get_msg(file_name)            
        self.in_edit_sender.setText(header[2])
        self.in_edit_sender_name.setText(header[1])
        self.in_edit_subject.setText(header[0])
        self.in_edit_html.insertPlainText(html)

    def eml_save(self):
        header, msg = [], ''
        header.append(self.in_edit_subject.text())
        header.append(self.in_edit_sender_name.text())
        header.append(self.in_edit_sender.text())
        header.append('test@email.com')
        annex_file = self.in_edit_annex.text().split(',')
        html = self.in_edit_html.toPlainText()

        if not any(header[:3]) or not html:
            return

        try:        
            msg = gm.gen_eml(header, html, annex_file) if self.cb_edit_annex.isChecked() else gm.gen_eml(header, html)

            file_path, _ = QFileDialog.getSaveFileName(self, '另存為...', './', 'Excel Files (*.eml)')
            with open(file_path, 'w') as outfile:
                gen = generator.Generator(outfile)
                gen.flatten(msg)
            
            QMessageBox.information(self, 'Success!', '儲存成功！', QMessageBox.Ok)
        except:
            QMessageBox.warning(self, 'Failed!', '儲存失敗！', QMessageBox.Ok)

    def eml_reset(self):
        items = self.get_items_from_layout(self.right_layout)
        for item in items:
            if type(item) == type(QLineEdit()):
                item.clear()
        
        self.cb_edit_annex.setChecked(False)
        self.in_edit_html.clear()

    def logs_change(self):
        if not self.data_logs or not self.in_logs_data.text():
            return

        self.data_temp_logs = []
        self.tbw_logs.setRowCount(0)
        
        # header = {'郵件類型':'type', '郵件主旨':'subject', '使用者群組':'user_group', '使用者信箱':'user_email'}
        condition = self.cmb_logs_choice.currentText()
        content = self.in_logs_data.text()

        row_num = len(self.data_logs)
        col_num = len(self.data_logs[0])

        # self.tbw_logs.setRowCount(row_num)  
        self.tbw_logs.setColumnCount(col_num)

        for i in range(row_num):
            switch = False
            if condition == 'date' and content in str(self.data_logs[i][condition]):
                switch = True
            elif self.data_logs[i][condition] == content:
                switch = True
                
            if switch:
                self.tbw_logs.insertRow(self.tbw_logs.rowCount())
                row_data = list(self.data_logs[i].values())
                self.data_temp_logs.append(self.data_logs[i])
                for j in range(col_num):
                    temp_data = row_data[j]
                    item = QTableWidgetItem(str(temp_data))
                    item.setForeground(QBrush(QColor(144, 182, 240)))
                    self.tbw_logs.setItem(self.tbw_logs.rowCount()-1, j, item)

    def logs_download(self):
        if self.data_temp_logs:
            try:
                file_path, _ = QFileDialog.getSaveFileName(self, '另存為...', './', 'Excel Files (*.xlsx)')
                if not file_path:
                    return

                df = DataFrame(self.data_temp_logs)
                df.to_excel(file_path, index=False)

                QMessageBox.information(self, 'Success!', '儲存成功！', QMessageBox.Ok)
            except:
                QMessageBox.warning(self, 'Failed!', '儲存失敗！', QMessageBox.Ok)
        else:
            QMessageBox.warning(self, "缺少資料", "請確認是否有資料可以下載", QMessageBox.Ok)

    def quit_act(self):
        # sender 是发送信号的对象
        sender = self.sender()
        print(sender.text() + '键被按下')
        qApp = QApplication.instance()
        qApp.quit()

def main():
    app = QApplication(sys.argv)
    gui = MailserverUi()
    gui.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()