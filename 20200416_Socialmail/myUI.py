# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'form.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

import sys
from PyQt5 import QtCore, QtGui, QtWidgets, QtWebEngineWidgets
from PyQt5.QtWidgets import QFileDialog, QWidget, QMainWindow, QApplication


import mylibrary.genmail as gm
from email import generator
from GenAndSendMail import insert_send_mail, get_data
from server.database import Database
from server.sendmail import Smtp
from server.client import Client

class Socialmails_UI(object):

    # def __init__(self):
    #     super().__init__()
    #     self.center()
    #     self.init_ui()
    #     self.setWindowTitle("社交郵件工程")

    #     self.main_widget = QWidget()
    #     self.setCentralWidget(self.main_widget)  

    #     # self.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
        
    #     self.setupUi(self.main_widget)

    def setupUi(self, myUI):
        myUI.setObjectName("myUI")
        myUI.resize(715, 471)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(myUI)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.main_tab = QtWidgets.QTabWidget(myUI)
        self.main_tab.setTabPosition(QtWidgets.QTabWidget.North)
        self.main_tab.setTabShape(QtWidgets.QTabWidget.Rounded)
        self.main_tab.setElideMode(QtCore.Qt.ElideRight)
        self.main_tab.setUsesScrollButtons(False)
        self.main_tab.setDocumentMode(True)
        self.main_tab.setTabsClosable(False)
        self.main_tab.setMovable(False)
        self.main_tab.setTabBarAutoHide(False)
        self.main_tab.setObjectName("main_tab")
        self.main_tab_1 = QtWidgets.QWidget()
        self.main_tab_1.setObjectName("main_tab_1")
        self.verticalLayout_6 = QtWidgets.QVBoxLayout(self.main_tab_1)
        self.verticalLayout_6.setObjectName("verticalLayout_6")

        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.verticalLayout_6.addItem(spacerItem)

        self.horizontalLayout_12 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_12.setObjectName("horizontalLayout_12")
        self.label = QtWidgets.QLabel(self.main_tab_1)
        self.label.setMinimumSize(QtCore.QSize(100, 0))
        self.label.setMaximumSize(QtCore.QSize(100, 40))
        self.label.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label.setObjectName("label")
        self.horizontalLayout_12.addWidget(self.label)
        self.label_4 = QtWidgets.QLabel(self.main_tab_1)
        self.label_4.setMinimumSize(QtCore.QSize(80, 0))
        self.label_4.setMaximumSize(QtCore.QSize(80, 16777215))
        self.label_4.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_12.addWidget(self.label_4)
        self.smtp_user = QtWidgets.QLineEdit(self.main_tab_1)
        self.smtp_user.setMinimumSize(QtCore.QSize(150, 0))
        self.smtp_user.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.smtp_user.setObjectName("smtp_user")
        self.horizontalLayout_12.addWidget(self.smtp_user)
        self.label_5 = QtWidgets.QLabel(self.main_tab_1)
        self.label_5.setMinimumSize(QtCore.QSize(100, 0))
        self.label_5.setMaximumSize(QtCore.QSize(100, 16777215))
        self.label_5.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_5.setObjectName("label_5")
        self.horizontalLayout_12.addWidget(self.label_5)
        self.smtp_pw = QtWidgets.QLineEdit(self.main_tab_1)
        self.smtp_pw.setMinimumSize(QtCore.QSize(150, 0))
        self.smtp_pw.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.smtp_pw.setObjectName("smtp_pw")
        self.horizontalLayout_12.addWidget(self.smtp_pw)
        self.verticalLayout_6.addLayout(self.horizontalLayout_12)

        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label_2 = QtWidgets.QLabel(self.main_tab_1)
        self.label_2.setMinimumSize(QtCore.QSize(190, 0))
        self.label_2.setMaximumSize(QtCore.QSize(190, 16777215))
        self.label_2.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout.addWidget(self.label_2)
        self.smtp_host = QtWidgets.QLineEdit(self.main_tab_1)
        self.smtp_host.setMinimumSize(QtCore.QSize(150, 0))
        self.smtp_host.setMaximumSize(QtCore.QSize(150, 16777215))
        self.smtp_host.setObjectName("smtp_host")
        self.horizontalLayout.addWidget(self.smtp_host)
        self.label_3 = QtWidgets.QLabel(self.main_tab_1)
        self.label_3.setMinimumSize(QtCore.QSize(50, 0))
        self.label_3.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout.addWidget(self.label_3)
        self.smtp_port = QtWidgets.QLineEdit(self.main_tab_1)
        self.smtp_port.setMinimumSize(QtCore.QSize(50, 0))
        self.smtp_port.setMaximumSize(QtCore.QSize(50, 16777215))
        self.smtp_port.setObjectName("smtp_port")
        self.horizontalLayout.addWidget(self.smtp_port)
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem)
        self.verticalLayout_6.addLayout(self.horizontalLayout)

        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.verticalLayout_6.addItem(spacerItem)

        self.verticalLayout_5 = QtWidgets.QVBoxLayout()
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.horizontalLayout_13 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_13.setObjectName("horizontalLayout_13")
        self.label_6 = QtWidgets.QLabel(self.main_tab_1)
        self.label_6.setMinimumSize(QtCore.QSize(100, 0))
        self.label_6.setMaximumSize(QtCore.QSize(100, 40))
        self.label_6.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_6.setObjectName("label_6")
        self.horizontalLayout_13.addWidget(self.label_6)
        self.label_8 = QtWidgets.QLabel(self.main_tab_1)
        self.label_8.setMinimumSize(QtCore.QSize(80, 0))
        self.label_8.setMaximumSize(QtCore.QSize(80, 16777215))
        self.label_8.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_8.setObjectName("label_8")
        self.horizontalLayout_13.addWidget(self.label_8)
        self.db_user = QtWidgets.QLineEdit(self.main_tab_1)
        self.db_user.setMinimumSize(QtCore.QSize(150, 0))
        self.db_user.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.db_user.setObjectName("db_user")
        self.horizontalLayout_13.addWidget(self.db_user)
        self.label_9 = QtWidgets.QLabel(self.main_tab_1)
        self.label_9.setMinimumSize(QtCore.QSize(100, 0))
        self.label_9.setMaximumSize(QtCore.QSize(100, 16777215))
        self.label_9.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_9.setObjectName("label_9")
        self.horizontalLayout_13.addWidget(self.label_9)
        self.db_pw = QtWidgets.QLineEdit(self.main_tab_1)
        self.db_pw.setMinimumSize(QtCore.QSize(175, 0))
        self.db_pw.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.db_pw.setObjectName("db_pw")
        self.horizontalLayout_13.addWidget(self.db_pw)
        self.verticalLayout_6.addLayout(self.horizontalLayout_13)

        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_7 = QtWidgets.QLabel(self.main_tab_1)
        self.label_7.setMinimumSize(QtCore.QSize(190, 0))
        self.label_7.setMaximumSize(QtCore.QSize(190, 16777215))
        self.label_7.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_7.setObjectName("label_7")
        self.horizontalLayout_2.addWidget(self.label_7)
        self.db_host = QtWidgets.QLineEdit(self.main_tab_1)
        self.db_host.setMinimumSize(QtCore.QSize(150, 0))
        self.db_host.setMaximumSize(QtCore.QSize(150, 16777215))
        self.db_host.setObjectName("db_host")
        self.horizontalLayout_2.addWidget(self.db_host)
        self.label_11 = QtWidgets.QLabel(self.main_tab_1)
        self.label_11.setMinimumSize(QtCore.QSize(50, 0))
        self.label_11.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_11.setObjectName("label_11")
        self.horizontalLayout_2.addWidget(self.label_11)
        self.db_port = QtWidgets.QLineEdit(self.main_tab_1)
        self.db_port.setMinimumSize(QtCore.QSize(50, 0))
        self.db_port.setMaximumSize(QtCore.QSize(50, 16777215))
        self.db_port.setObjectName("db_port")
        self.horizontalLayout_2.addWidget(self.db_port)
        self.label_10 = QtWidgets.QLabel(self.main_tab_1)
        self.label_10.setMinimumSize(QtCore.QSize(100, 0))
        self.label_10.setMaximumSize(QtCore.QSize(100, 16777215))
        self.label_10.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_10.setObjectName("label_10")
        self.horizontalLayout_2.addWidget(self.label_10)
        self.db_db = QtWidgets.QLineEdit(self.main_tab_1)
        self.db_db.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.db_db.setObjectName("db_db")
        self.horizontalLayout_2.addWidget(self.db_db)
        self.verticalLayout_6.addLayout(self.horizontalLayout_2)
        
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.verticalLayout_6.addItem(spacerItem)

        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.label_12 = QtWidgets.QLabel(self.main_tab_1)
        self.label_12.setMinimumSize(QtCore.QSize(100, 0))
        self.label_12.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.label_12.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_12.setObjectName("label_12")
        self.horizontalLayout_3.addWidget(self.label_12)
        self.eml = QtWidgets.QLineEdit(self.main_tab_1)
        self.eml.setObjectName("eml")
        self.horizontalLayout_3.addWidget(self.eml)
        self.eml_btn = QtWidgets.QPushButton(self.main_tab_1)
        self.eml_btn.setMinimumSize(QtCore.QSize(120, 0))
        self.eml_btn.setObjectName("eml_btn")
        self.horizontalLayout_3.addWidget(self.eml_btn)
        self.label_21 = QtWidgets.QLabel(self.main_tab_1)
        self.label_21.setMinimumSize(QtCore.QSize(50, 0))
        self.label_21.setMaximumSize(QtCore.QSize(50, 16777215))
        self.label_21.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_21.setObjectName("label_21")
        self.horizontalLayout_3.addWidget(self.label_21)
        self.eml_type = QtWidgets.QLineEdit(self.main_tab_1)
        self.eml_type.setMinimumSize(QtCore.QSize(80, 0))
        self.eml_type.setMaximumSize(QtCore.QSize(120, 16777215))
        self.eml_type.setObjectName("eml_type")
        self.horizontalLayout_3.addWidget(self.eml_type)
        self.verticalLayout_6.addLayout(self.horizontalLayout_3)

        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_13 = QtWidgets.QLabel(self.main_tab_1)
        self.label_13.setMinimumSize(QtCore.QSize(100, 0))
        self.label_13.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.label_13.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_13.setObjectName("label_13")
        self.horizontalLayout_4.addWidget(self.label_13)
        self.mails = QtWidgets.QLineEdit(self.main_tab_1)
        self.mails.setObjectName("mails")
        self.horizontalLayout_4.addWidget(self.mails)
        self.mails_btn = QtWidgets.QPushButton(self.main_tab_1)
        self.mails_btn.setMinimumSize(QtCore.QSize(120, 0))
        self.mails_btn.setObjectName("mails_btn")
        self.horizontalLayout_4.addWidget(self.mails_btn)
        self.label_22 = QtWidgets.QLabel(self.main_tab_1)
        self.label_22.setMinimumSize(QtCore.QSize(50, 0))
        self.label_22.setMaximumSize(QtCore.QSize(50, 16777215))
        self.label_22.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_22.setObjectName("label_22")
        self.horizontalLayout_4.addWidget(self.label_22)
        self.mails_group = QtWidgets.QLineEdit(self.main_tab_1)
        self.mails_group.setMinimumSize(QtCore.QSize(80, 0))
        self.mails_group.setMaximumSize(QtCore.QSize(120, 16777215))
        self.mails_group.setObjectName("mails_group")
        self.horizontalLayout_4.addWidget(self.mails_group)
        self.verticalLayout_6.addLayout(self.horizontalLayout_4)

        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.label_14 = QtWidgets.QLabel(self.main_tab_1)
        self.label_14.setMinimumSize(QtCore.QSize(100, 0))
        self.label_14.setMaximumSize(QtCore.QSize(16777215, 40))
        self.label_14.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_14.setObjectName("label_14")
        self.horizontalLayout_5.addWidget(self.label_14)
        self.datetime_edit = QtWidgets.QDateTimeEdit(self.main_tab_1)
        self.datetime_edit.setObjectName("datetime_edit")
        self.datetime_edit.setCalendarPopup(True)
        self.datetime_edit.setDateTime(QtCore.QDateTime.currentDateTime())
        self.datetime_edit.setDisplayFormat('yyyy-MM-dd hh:mm')
        self.horizontalLayout_5.addWidget(self.datetime_edit)
        self.add_cb = QtWidgets.QCheckBox(self.main_tab_1)
        self.add_cb.setObjectName("add_cb")
        self.horizontalLayout_5.addWidget(self.add_cb)
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_5.addItem(spacerItem1)
        self.run_btn = QtWidgets.QPushButton(self.main_tab_1)
        self.run_btn.setMinimumSize(QtCore.QSize(120, 0))
        self.run_btn.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.run_btn.setObjectName("run_btn")
        self.horizontalLayout_5.addWidget(self.run_btn)
        self.verticalLayout_6.addLayout(self.horizontalLayout_5)
        spacerItem2 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_6.addItem(spacerItem2)

        self.main_tab.addTab(self.main_tab_1, "")
        self.main_tab_2 = QtWidgets.QWidget()
        self.main_tab_2.setObjectName("main_tab_2")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.main_tab_2)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.horizontalLayout_11 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_11.setObjectName("horizontalLayout_11")
        self.label_15 = QtWidgets.QLabel(self.main_tab_2)
        self.label_15.setMinimumSize(QtCore.QSize(80, 0))
        self.label_15.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_15.setObjectName("label_15")
        self.horizontalLayout_11.addWidget(self.label_15)
        self.recipient = QtWidgets.QLineEdit(self.main_tab_2)
        self.recipient.setObjectName("recipient")
        self.horizontalLayout_11.addWidget(self.recipient)
        self.verticalLayout_3.addLayout(self.horizontalLayout_11)
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        self.label_16 = QtWidgets.QLabel(self.main_tab_2)
        self.label_16.setMinimumSize(QtCore.QSize(80, 0))
        self.label_16.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_16.setObjectName("label_16")
        self.horizontalLayout_10.addWidget(self.label_16)
        self.recipient_name = QtWidgets.QLineEdit(self.main_tab_2)
        self.recipient_name.setObjectName("recipient_name")
        self.horizontalLayout_10.addWidget(self.recipient_name)
        self.verticalLayout_3.addLayout(self.horizontalLayout_10)
        self.horizontalLayout_9 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_9.setObjectName("horizontalLayout_9")
        self.label_17 = QtWidgets.QLabel(self.main_tab_2)
        self.label_17.setMinimumSize(QtCore.QSize(80, 0))
        self.label_17.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_17.setObjectName("label_17")
        self.horizontalLayout_9.addWidget(self.label_17)
        self.annex_checkBox = QtWidgets.QCheckBox(self.main_tab_2)
        self.annex_checkBox.setObjectName("annex_checkBox")
        self.horizontalLayout_9.addWidget(self.annex_checkBox)
        spacerItem3 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_9.addItem(spacerItem3)
        self.verticalLayout_3.addLayout(self.horizontalLayout_9)
        self.horizontalLayout_8 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_8.setObjectName("horizontalLayout_8")
        self.label_18 = QtWidgets.QLabel(self.main_tab_2)
        self.label_18.setMinimumSize(QtCore.QSize(80, 0))
        self.label_18.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_18.setObjectName("label_18")
        self.horizontalLayout_8.addWidget(self.label_18)
        self.annex_file = QtWidgets.QLineEdit(self.main_tab_2)
        self.annex_file.setObjectName("annex_file")
        self.horizontalLayout_8.addWidget(self.annex_file)
        self.annex_btn = QtWidgets.QPushButton(self.main_tab_2)
        self.annex_btn.setObjectName("annex_btn")
        self.horizontalLayout_8.addWidget(self.annex_btn)
        self.verticalLayout_3.addLayout(self.horizontalLayout_8)
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.label_19 = QtWidgets.QLabel(self.main_tab_2)
        self.label_19.setMinimumSize(QtCore.QSize(80, 0))
        self.label_19.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_19.setObjectName("label_19")
        self.horizontalLayout_7.addWidget(self.label_19)
        self.subject = QtWidgets.QLineEdit(self.main_tab_2)
        self.subject.setObjectName("subject")
        self.horizontalLayout_7.addWidget(self.subject)
        self.verticalLayout_3.addLayout(self.horizontalLayout_7)
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.label_20 = QtWidgets.QLabel(self.main_tab_2)
        self.label_20.setMinimumSize(QtCore.QSize(80, 0))
        self.label_20.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_20.setObjectName("label_20")
        self.verticalLayout.addWidget(self.label_20)
        spacerItem4 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem4)
        self.horizontalLayout_6.addLayout(self.verticalLayout)
        
        self.mail_tab = QtWidgets.QTabWidget(self.main_tab_2)
        self.mail_tab.setDocumentMode(True)
        self.mail_tab.setObjectName("mail_tab")
        self.mail_tab_1 = QtWidgets.QWidget()
        self.mail_tab_1.setAutoFillBackground(False)
        self.mail_tab_1.setObjectName("mail_tab_1")
        self.verticalLayout_7 = QtWidgets.QVBoxLayout(self.mail_tab_1)
        self.verticalLayout_7.setSizeConstraint(QtWidgets.QLayout.SetMaximumSize)
        self.verticalLayout_7.setObjectName("verticalLayout_7")
        self.textEdit = QtWidgets.QTextEdit(self.mail_tab_1)
        self.textEdit.setObjectName("textEdit")
        self.verticalLayout_7.addWidget(self.textEdit)
        self.mail_tab.addTab(self.mail_tab_1, "")
        self.mail_tab_2 = QtWidgets.QWidget()
        self.mail_tab_2.setObjectName("mail_tab_2")
        self.verticalLayout_8 = QtWidgets.QVBoxLayout(self.mail_tab_2)
        self.verticalLayout_8.setObjectName("verticalLayout_8")
        self.webEngineView = QtWebEngineWidgets.QWebEngineView(self.mail_tab_2)
        self.webEngineView.setObjectName('webEngineView')
        self.verticalLayout_8.addWidget(self.webEngineView)
        self.mail_tab.addTab(self.mail_tab_2, "")
        self.horizontalLayout_6.addWidget(self.mail_tab)
        self.verticalLayout_3.addLayout(self.horizontalLayout_6)
        self.horizontalLayout_14 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_14.setObjectName("horizontalLayout_14")
        spacerItem5 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_14.addItem(spacerItem5)
        self.reset_eml_btn = QtWidgets.QPushButton(self.main_tab_2)
        self.reset_eml_btn.setObjectName('reset_eml_btn')
        self.horizontalLayout_14.addWidget(self.reset_eml_btn)
        self.read_eml_btn = QtWidgets.QPushButton(self.main_tab_2)
        self.read_eml_btn.setObjectName("read_eml_btn")
        self.horizontalLayout_14.addWidget(self.read_eml_btn)
        self.save_eml_btn = QtWidgets.QPushButton(self.main_tab_2)
        self.save_eml_btn.setObjectName("save_eml_btn")
        self.horizontalLayout_14.addWidget(self.save_eml_btn)
        self.verticalLayout_3.addLayout(self.horizontalLayout_14)
        self.main_tab.addTab(self.main_tab_2, "")
        self.verticalLayout_2.addWidget(self.main_tab)

        self.set_signal(myUI)
        self.retranslateUi(myUI)
        self.main_tab.setCurrentIndex(1)
        self.mail_tab.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(myUI)

        self.current_eml_file = None

    def retranslateUi(self, myUI):
        _translate = QtCore.QCoreApplication.translate
        myUI.setWindowTitle(_translate("myUI", "myUI"))
        self.label.setText(_translate("myUI", "郵件伺服器 : "))
        self.label_2.setText(_translate("myUI", "host : "))
        self.label_3.setText(_translate("myUI", "port "))
        self.label_4.setText(_translate("myUI", "帳號 : "))
        self.label_5.setText(_translate("myUI", "密碼 : "))
        self.label_6.setText(_translate("myUI", "資料庫 : "))
        self.label_7.setText(_translate("myUI", "host :"))
        self.label_11.setText(_translate("myUI", "port "))
        self.label_10.setText(_translate("myUI", "Database :"))
        self.label_8.setText(_translate("myUI", "帳號 :"))
        self.label_9.setText(_translate("myUI", "密碼 :"))
        self.label_12.setText(_translate("myUI", "信件模板 :"))
        self.eml_btn.setText(_translate("myUI", "瀏覽"))
        self.label_21.setText(_translate("myUI", "類型 :"))
        self.label_13.setText(_translate("myUI", "收信人資料 :"))
        self.mails_btn.setText(_translate("myUI", "瀏覽"))
        self.label_22.setText(_translate("myUI", "群組 :"))
        self.label_14.setText(_translate("myUI", "設定排程 :"))
        self.add_cb.setText(_translate("myUI", "新增"))
        self.run_btn.setText(_translate("myUI", "開始執行"))
        self.main_tab.setTabText(self.main_tab.indexOf(self.main_tab_1), _translate("myUI", "寄送郵件"))
        self.label_15.setText(_translate("myUI", "寄件人 :"))
        self.label_16.setText(_translate("myUI", "寄件人名稱 :"))
        self.label_17.setText(_translate("myUI", "加入附件 :"))
        self.annex_checkBox.setText(_translate("myUI", "是"))
        self.label_18.setText(_translate("myUI", "上傳附件 :"))
        self.annex_btn.setText(_translate("myUI", "瀏覽"))
        self.label_19.setText(_translate("myUI", "主旨 :"))
        self.label_20.setText(_translate("myUI", "內容 :"))
        self.mail_tab.setTabText(self.mail_tab.indexOf(self.mail_tab_1), _translate("myUI", "HTML"))
        self.mail_tab.setTabText(self.mail_tab.indexOf(self.mail_tab_2), _translate("myUI", "Web"))
        self.reset_eml_btn.setText(_translate("myUI", "重置"))
        self.read_eml_btn.setText(_translate("myUI", "讀取"))
        self.save_eml_btn.setText(_translate("myUI", "儲存"))
        self.main_tab.setTabText(self.main_tab.indexOf(self.main_tab_2), _translate("myUI", "編輯信件"))

    def set_signal(self, myUI):
        self.eml_btn.clicked.connect(self.open_file)
        self.mails_btn.clicked.connect(self.open_file)
        self.run_btn.clicked.connect(self.send_mail)

        self.annex_btn.clicked.connect(self.open_file)
        self.mail_tab.currentChanged.connect(self.print_html)
        self.read_eml_btn.clicked.connect(self.open_eml)
        self.save_eml_btn.clicked.connect(self.save_eml)
        self.reset_eml_btn.clicked.connect(self.reset_eml)

    def open_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", "All Files (*)")
        if '.zip' in file_name:
            org_files = self.annex_file.text()
            all_files = org_files + ',' + file_name if org_files else file_name
            self.annex_file.setText(all_files)
        elif '.jpg' in file_name or '.png' in file_name:
            org_files = self.annex_file.text()
            all_files = org_files + ',' + file_name if org_files else file_name
            self.annex_file.setText(all_files)
        elif '.eml' in file_name:
            self.eml.setText(file_name)
        elif '.xlsx' in file_name:
            self.mails.setText(file_name)
            
    def open_eml(self, obj):
        file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", "Eml Files (*.eml)")
        obj.setText(file_name)

    def open_excel(self, obj):
        file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", "Excel Files (*.xlsx)")
        obj.setText(file_name)

    def open_path(self, obj):
        directory = QFileDialog.getExistingDirectory(self, "選取資料夾", "./")
        obj.setText(directory+'/')

    def send_mail(self):
        smtp_host = self.smtp_host.text() if self.smtp_host.text() else 'localhost'
        smtp_port = self.smtp_port.text() if self.smtp_port.text() else 465
        smtp_user = self.smtp_user.text() if self.smtp_user.text() else 'yucc@yumail.myvnc.com'
        smtp_pw = self.smtp_pw.text() if self.smtp_pw.text() else 'yucc1234'
        db_host = self.db_host.text() if self.db_host.text() else 'localhost'
        db_port = self.db_port.text() if self.db_port.text() else 3306
        db_user = self.db_user.text() if self.db_user.text() else 'socialmails'
        db_pw = self.db_pw.text() if self.db_pw.text() else 'socialmails123'
        db_db = self.db_db.text() if self.db_db.text() else 'socialmails'
        eml_file = self.eml.text()
        eml_type = self.eml_type.text()
        mails_excel = self.mails.text()
        mails_group = self.mails_group.text()

        if self.add_cb.isChecked():
            my_time = self.datetime_edit.text() + ':00'
            client = Client()
            client.send(eml_type, eml_file, mails_group, mails_excel, my_time)

        else:
            db = Database(db_host, db_port, db_user, db_pw, db_db)
            sm = Smtp(smtp_host, smtp_port, smtp_user, smtp_pw)
            
            insert_send_mail(eml_type, eml_file, mails_group, mails_excel, sm, db)

    def print_html(self, index):
        if index:
            self.webEngineView.setHtml(self.textEdit.toPlainText())
    
    def open_eml(self):
        self.textEdit.clear()
        # self.annex_checkBox.setChecked(False)
        # self.annex_checkBox.setDisabled(True)
        # self.annex_file.setDisabled(True)
        try:
            file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", "All Files (*)")
            header, html = gm.get_msg(file_name)            
            self.recipient.setText(header[2])
            self.recipient_name.setText(header[1])
            self.subject.setText(header[0])
            self.textEdit.insertPlainText(html)
            
            self.current_eml_file = file_name
        except:
            pass

    def save_eml(self):
        header, msg = [], ''
        header.append(self.subject.text())
        header.append(self.recipient_name.text())
        header.append(self.recipient.text())
        header.append('test@email.com')
        annex_file = self.annex_file.text().split(',')
        html = self.textEdit.toPlainText()
        
        if self.current_eml_file:
            msg = ''
            try:
                msg = gm.update_eml(self.current_eml_file, header[3], html, [header[1], header[2]])
                if self.annex_checkBox.checkState():
                    gm.msg_add_annex(msg, annex_file)
            except:
                msg = gm.gen_eml(header, html, annex_file)

            with open(self.current_eml_file+'.new', 'w') as outfile:
                gen = generator.Generator(outfile)
                gen.flatten(msg)
        else:
            if self.annex_checkBox.checkState():
                msg = gm.gen_eml(header, html, annex_file)
            else:
                msg = gm.gen_eml(header, html)
            
            directory = QFileDialog.getExistingDirectory(self, "選取資料夾", "./")
            with open(directory+'/'+header[0]+'.eml', 'w') as outfile:
                gen = generator.Generator(outfile)
                gen.flatten(msg)
                
    def reset_eml(self):
        self.recipient.clear()
        self.recipient_name.clear()
        self.subject.clear()
        self.textEdit.clear()
        self.annex_checkBox.setEnabled(True)
        self.annex_file.setEnabled(True)
        self.annex_file.clear()
        
        self.current_eml_file = None

class MainWindow(QMainWindow, Socialmails_UI):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        central_widget = QWidget()
        self.setCentralWidget(central_widget) # new central widget    
        self.setupUi(central_widget)

# def main():
#     app = QApplication(sys.argv)
#     gui = Socialmails_UI()
#     gui.show()
#     sys.exit(app.exec_())

if __name__ == "__main__":
    # main()
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())