#!/usr/bin/env python
# coding: utf-8

# In[1]:


from ExcelToWordCostum import ExcelToWord

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox

class Ui_ExcelToWord(object):
    def setupUi(self, ExcelToWord):
        ExcelToWord.setObjectName("ExcelToWord")
        ExcelToWord.resize(560, 420)
        ExcelToWord.setMinimumSize(QtCore.QSize(560, 420))
        ExcelToWord.setMaximumSize(QtCore.QSize(560, 420))
        self.verticalLayout = QtWidgets.QVBoxLayout(ExcelToWord)
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.Excel = QtWidgets.QLabel(ExcelToWord)
        self.Excel.setObjectName("Excel")
        self.horizontalLayout.addWidget(self.Excel)
        self.Excel_in = QtWidgets.QLineEdit(ExcelToWord)
        self.Excel_in.setObjectName("Excel_in")
        self.horizontalLayout.addWidget(self.Excel_in)
        self.Sheet = QtWidgets.QLabel(ExcelToWord)
        self.Sheet.setObjectName("Sheet")
        self.horizontalLayout.addWidget(self.Sheet)
        self.Sheet_in = QtWidgets.QLineEdit(ExcelToWord)
        self.Sheet_in.setMaximumSize(QtCore.QSize(80, 16777215))
        self.Sheet_in.setObjectName("Sheet_in")
        self.horizontalLayout.addWidget(self.Sheet_in)
        self.Excel_btn = QtWidgets.QPushButton(ExcelToWord)
        self.Excel_btn.setMinimumSize(QtCore.QSize(100, 0))
        self.Excel_btn.setObjectName("Excel_btn")
        self.horizontalLayout.addWidget(self.Excel_btn)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.Word = QtWidgets.QLabel(ExcelToWord)
        self.Word.setObjectName("Word")
        self.horizontalLayout_2.addWidget(self.Word)
        self.Word_in = QtWidgets.QLineEdit(ExcelToWord)
        self.Word_in.setObjectName("Word_in")
        self.horizontalLayout_2.addWidget(self.Word_in)
        self.Word_btn = QtWidgets.QPushButton(ExcelToWord)
        self.Word_btn.setMinimumSize(QtCore.QSize(100, 0))
        self.Word_btn.setObjectName("Word_btn")
        self.horizontalLayout_2.addWidget(self.Word_btn)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.Year = QtWidgets.QLabel(ExcelToWord)
        self.Year.setObjectName("Year")
        self.horizontalLayout_3.addWidget(self.Year)
        self.Year_in = QtWidgets.QLineEdit(ExcelToWord)
        self.Year_in.setMaximumSize(QtCore.QSize(80, 16777215))
        self.Year_in.setObjectName("Year_in")
        self.horizontalLayout_3.addWidget(self.Year_in)
        self.Month = QtWidgets.QLabel(ExcelToWord)
        self.Month.setObjectName("Month")
        self.horizontalLayout_3.addWidget(self.Month)
        self.Month_in = QtWidgets.QLineEdit(ExcelToWord)
        self.Month_in.setMaximumSize(QtCore.QSize(60, 16777215))
        self.Month_in.setObjectName("Month_in")
        self.horizontalLayout_3.addWidget(self.Month_in)
        self.Day = QtWidgets.QLabel(ExcelToWord)
        self.Day.setObjectName("Day")
        self.horizontalLayout_3.addWidget(self.Day)
        self.Day_in = QtWidgets.QLineEdit(ExcelToWord)
        self.Day_in.setMaximumSize(QtCore.QSize(60, 16777215))
        self.Day_in.setObjectName("Day_in")
        self.horizontalLayout_3.addWidget(self.Day_in)
        self.IP = QtWidgets.QLabel(ExcelToWord)
        self.IP.setObjectName("IP")
        self.horizontalLayout_3.addWidget(self.IP)
        self.IP_in = QtWidgets.QLineEdit(ExcelToWord)
        self.IP_in.setText("")
        self.IP_in.setObjectName("IP_in")
        self.horizontalLayout_3.addWidget(self.IP_in)
        self.verticalLayout.addLayout(self.horizontalLayout_3)
        self.textEdit = QtWidgets.QTextEdit(ExcelToWord)
        self.textEdit.setObjectName("textEdit")
        self.verticalLayout.addWidget(self.textEdit)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.Save = QtWidgets.QLabel(ExcelToWord)
        self.Save.setObjectName("Save")
        self.horizontalLayout_4.addWidget(self.Save)
        self.Save_in = QtWidgets.QLineEdit(ExcelToWord)
        self.Save_in.setObjectName("Save_in")
        self.horizontalLayout_4.addWidget(self.Save_in)
        self.Savefile = QtWidgets.QLabel(ExcelToWord)
        self.Savefile.setObjectName("Savefile")
        self.horizontalLayout_4.addWidget(self.Savefile)
        self.Savefile_in = QtWidgets.QLineEdit(ExcelToWord)
        self.Savefile_in.setMaximumSize(QtCore.QSize(100, 16777215))
        self.Savefile_in.setObjectName("Savefile_in")
        self.horizontalLayout_4.addWidget(self.Savefile_in)
        self.Save_btn = QtWidgets.QPushButton(ExcelToWord)
        self.Save_btn.setMinimumSize(QtCore.QSize(100, 0))
        self.Save_btn.setObjectName("Save_btn")
        self.horizontalLayout_4.addWidget(self.Save_btn)
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_6.addItem(spacerItem)
        self.Clear_btn = QtWidgets.QPushButton(ExcelToWord)
        self.Clear_btn.setMinimumSize(QtCore.QSize(100, 0))
        self.Clear_btn.setMaximumSize(QtCore.QSize(100, 16777215))
        self.Clear_btn.setObjectName("Clear_btn")
        self.horizontalLayout_6.addWidget(self.Clear_btn)
        self.Excute_btn = QtWidgets.QPushButton(ExcelToWord)
        self.Excute_btn.setMinimumSize(QtCore.QSize(100, 0))
        self.Excute_btn.setObjectName("Excute_btn")
        self.horizontalLayout_6.addWidget(self.Excute_btn)
        self.verticalLayout.addLayout(self.horizontalLayout_6)

        self.retranslateUi(ExcelToWord)
        self.Excel_btn.clicked.connect(self.open_file)
        self.Word_btn.clicked.connect(self.open_file)
        self.Save_btn.clicked.connect(self.save_path)
        self.Excute_btn.clicked.connect(self.input_data)
        
        self.Clear_btn.clicked.connect(self.Excel_in.clear)
        self.Clear_btn.clicked.connect(self.Sheet_in.clear)
        self.Clear_btn.clicked.connect(self.Word_in.clear)
        self.Clear_btn.clicked.connect(self.Year_in.clear)
        self.Clear_btn.clicked.connect(self.Month_in.clear)
        self.Clear_btn.clicked.connect(self.Day_in.clear)
        self.Clear_btn.clicked.connect(self.IP_in.clear)
        self.Clear_btn.clicked.connect(self.textEdit.clear)
        self.Clear_btn.clicked.connect(self.Save_in.clear)
        self.Clear_btn.clicked.connect(self.Savefile_in.clear)
        
        QtCore.QMetaObject.connectSlotsByName(ExcelToWord)

    def retranslateUi(self, ExcelToWord):
        _translate = QtCore.QCoreApplication.translate
        ExcelToWord.setWindowTitle(_translate("ExcelToWord", "Excel To Word v.1.0.1"))
        self.Excel.setText(_translate("ExcelToWord", "Excel"))
        self.Sheet.setText(_translate("ExcelToWord", "Sheet"))
        self.Excel_btn.setText(_translate("ExcelToWord", "選取 Excel"))
        self.Word.setText(_translate("ExcelToWord", "Word"))
        self.Word_btn.setText(_translate("ExcelToWord", "選取 Word"))
        self.Year.setText(_translate("ExcelToWord", "Year"))
        self.Month.setText(_translate("ExcelToWord", "Month"))
        self.Day.setText(_translate("ExcelToWord", "Day"))
        self.IP.setText(_translate("ExcelToWord", "指定 IP (option)"))
        self.textEdit.setText(_translate("ExcelToWord", "初始化成功....."))
        self.Save.setText(_translate("ExcelToWord", "儲存至："))
        self.Savefile.setText(_translate("ExcelToWord", "檔案名"))
        self.Save_btn.setText(_translate("ExcelToWord", "儲存路徑"))
        self.Clear_btn.setText(_translate("ExcelToWord", "清除"))
        self.Excute_btn.setText(_translate("ExcelToWord", "執行"))
        
    def open_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", "All Files (*)")  
        if "xlsx" in file_name:
            self.Excel_in.setText(file_name)
        elif "docx" in file_name:
            self.Word_in.setText(file_name)
            
    def save_path(self):
        directory = QFileDialog.getExistingDirectory(self, "選取資料夾", "./")  
        self.Save_in.setText(directory)
        
    def input_data(self):
        excel = self.Excel_in.text()
        sheet = self.Sheet_in.text()
        word = self.Word_in.text()
        date = [self.Year_in.text(), self.Month_in.text(), self.Day_in.text()]
        ips = []
        save = self.Save_in.text() + "/" + self.Savefile_in.text()
        
        if not excel or not sheet or not word or not date or not save:
            QMessageBox.warning(self, "缺少資料", "請確認必要資料是否填入", QMessageBox.Ok)
            
        else:
            if self.IP_in.text():
                ips.append(self.IP_in.text())

            self.textEdit.append("\nExcel Path : %s" % excel)
            self.textEdit.append("Sheet : %s" % sheet)
            self.textEdit.append("Word Path : %s" % word)
            self.textEdit.append("Date : %s/%s/%s" % (date[0], date[1], date[2]))
            self.textEdit.append("指定 IP : %s" % ips)
            self.textEdit.append("Save file Path : %s" % save)

            ExcelToWord(excel, sheet, word, date, ips, save)
        
    
        


# In[ ]:


import sys
from PyQt5.QtWidgets import QWidget, QMainWindow, QApplication

class MainWindow(QMainWindow, Ui_ExcelToWord):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        central_widget = QWidget()
        self.setCentralWidget(central_widget) # new central widget    
        self.setupUi(central_widget)
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


# In[ ]:




