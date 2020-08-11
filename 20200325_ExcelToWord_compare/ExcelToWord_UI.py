# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'ExcelToWord_cp.ui'
#
# Created by: PyQt5 UI code generator 5.14.1
#
# WARNING! All changes made in this file will be lost!

import sys
import time

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QWidget, QMainWindow, QApplication
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from ExcelToWord_Copare import ExcelToWord_compare
from ExcelToWord_Copare import ExcelToExcel_compare

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
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.Excel_2 = QtWidgets.QLabel(ExcelToWord)
        self.Excel_2.setObjectName("Excel_2")
        self.horizontalLayout_3.addWidget(self.Excel_2)
        self.Excel_in_2 = QtWidgets.QLineEdit(ExcelToWord)
        self.Excel_in_2.setObjectName("Excel_in_2")
        self.horizontalLayout_3.addWidget(self.Excel_in_2)
        self.Sheet_2 = QtWidgets.QLabel(ExcelToWord)
        self.Sheet_2.setObjectName("Sheet_2")
        self.horizontalLayout_3.addWidget(self.Sheet_2)
        self.Sheet_in_2 = QtWidgets.QLineEdit(ExcelToWord)
        self.Sheet_in_2.setMaximumSize(QtCore.QSize(80, 16777215))
        self.Sheet_in_2.setObjectName("Sheet_in_2")
        self.horizontalLayout_3.addWidget(self.Sheet_in_2)
        self.Excel_btn_2 = QtWidgets.QPushButton(ExcelToWord)
        self.Excel_btn_2.setMinimumSize(QtCore.QSize(100, 0))
        self.Excel_btn_2.setObjectName("Excel_btn_2")
        self.horizontalLayout_3.addWidget(self.Excel_btn_2)
        self.verticalLayout.addLayout(self.horizontalLayout_3)
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
        self.Rb_excel = QtWidgets.QRadioButton(ExcelToWord)
        self.Rb_excel.setObjectName("Rb_excel")
        self.horizontalLayout_6.addWidget(self.Rb_excel)
        self.Rb_word = QtWidgets.QRadioButton(ExcelToWord)
        self.Rb_word.setObjectName("Rb_word")
        self.horizontalLayout_6.addWidget(self.Rb_word)
        self.Info_btn = QtWidgets.QPushButton(ExcelToWord)
        self.Info_btn.setMinimumSize(QtCore.QSize(100, 0))
        self.Info_btn.setObjectName("Info_btn")
        self.horizontalLayout_6.addWidget(self.Info_btn)
        self.Clear_btn = QtWidgets.QPushButton(ExcelToWord)
        self.Clear_btn.setMinimumSize(QtCore.QSize(100, 0))
        self.Clear_btn.setObjectName("Clear_btn")
        self.horizontalLayout_6.addWidget(self.Clear_btn)
        self.Excute_btn = QtWidgets.QPushButton(ExcelToWord)
        self.Excute_btn.setMinimumSize(QtCore.QSize(100, 0))
        self.Excute_btn.setObjectName("Excute_btn")
        self.horizontalLayout_6.addWidget(self.Excute_btn)
        self.verticalLayout.addLayout(self.horizontalLayout_6)

        self.retranslateUi(ExcelToWord)
        self.clickedconnect()
        QtCore.QMetaObject.connectSlotsByName(ExcelToWord)

        
    def retranslateUi(self, ExcelToWord):
        _translate = QtCore.QCoreApplication.translate
        ExcelToWord.setWindowTitle(_translate("ExcelToWord", "弱點掃描自動報表"))
        self.Excel.setText(_translate("ExcelToWord", "初掃 Excel"))
        self.Sheet.setText(_translate("ExcelToWord", "Sheet"))
        self.Excel_btn.setText(_translate("ExcelToWord", "選取 Excel"))
        self.Excel_2.setText(_translate("ExcelToWord", "複掃 Excel"))
        self.Sheet_2.setText(_translate("ExcelToWord", "Sheet"))
        self.Excel_btn_2.setText(_translate("ExcelToWord", "選取 Excel"))
        self.Word.setText(_translate("ExcelToWord", "模板 Word"))
        self.Word_btn.setText(_translate("ExcelToWord", "選取 Word"))
        self.Save.setText(_translate("ExcelToWord", "儲存至："))
        self.Savefile.setText(_translate("ExcelToWord", "檔案名"))
        self.Save_btn.setText(_translate("ExcelToWord", "儲存路徑"))
        self.Rb_excel.setText(_translate("ExcelToWord", "匯出 Excel"))
        self.Rb_word.setText(_translate("ExcelToWord", "匯出 Word"))
        self.Info_btn.setText(_translate("ExcelToWord", "使用說明"))
        self.Clear_btn.setText(_translate("ExcelToWord", "清除"))
        self.Excute_btn.setText(_translate("ExcelToWord", "執行"))

    def clickedconnect(self):
        self.Excel_btn.clicked.connect(self.open_excel_1)
        self.Excel_btn_2.clicked.connect(self.open_excel_2)
        self.Word_btn.clicked.connect(self.open_file)
        self.Save_btn.clicked.connect(self.save_path)
        self.Excute_btn.clicked.connect(self.input_data)
        self.Info_btn.clicked.connect(self.selectInfo)

        self.Clear_btn.clicked.connect(self.Excel_in.clear)
        self.Clear_btn.clicked.connect(self.Excel_in_2.clear)
        self.Clear_btn.clicked.connect(self.Sheet_in.clear)
        self.Clear_btn.clicked.connect(self.Sheet_in_2.clear)
        self.Clear_btn.clicked.connect(self.Word_in.clear)
        self.Clear_btn.clicked.connect(self.textEdit.clear)
        self.Clear_btn.clicked.connect(self.Save_in.clear)
        self.Clear_btn.clicked.connect(self.Savefile_in.clear)

    def selectInfo(self):
        title_text = '可放進 Word 中的標籤：\n\n'
        table_text = '表單類： 需插入 1*1 表格並填入以下標籤\n\n1. 初複掃弱點數量比較：riskCntCompare\n\n'

        dialog = QMessageBox(self)
        dialog.setInformativeText(title_text+table_text)
        dialog.setIcon(QMessageBox.Information)
        dialog.show()

    def open_excel_1(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", "All Files (*)")
        self.Excel_in.setText(file_name)

    def open_excel_2(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", "All Files (*)")
        self.Excel_in_2.setText(file_name)

    def open_file(self, in_lineedit):
        file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", "All Files (*)")
        self.Word_in.setText(file_name)

    def save_path(self):
        directory = QFileDialog.getExistingDirectory(self, "選取資料夾", "./")
        self.Save_in.setText(directory)
        
    def input_data(self):
        excel = self.Excel_in.text()
        excel2 = self.Excel_in_2.text()
        sheet = self.Sheet_in.text()
        sheet2 = self.Sheet_in_2.text()
        word = self.Word_in.text()
        save = self.Save_in.text() + "/" + self.Savefile_in.text()

        if not excel or not excel2 or not sheet or not sheet2 or not word or not save:
            QMessageBox.warning(self, "缺少資料", "請確認必要資料是否填入", QMessageBox.Ok)

        else:
            self.textEdit.append("初掃 Excel : %s" % excel)
            self.textEdit.append("複掃 Excel : %s" % excel2)          
            self.textEdit.append("模板  Word : %s" % word)
            self.textEdit.append("Save file Path : %s" % save)
            self.textEdit.append("執行中...\n")
            self.textEdit.moveCursor(QtGui.QTextCursor.End)
            QApplication.processEvents()
            time.sleep(1)

            if self.Rb_word.isChecked():
                ExcelToWord_compare(word, excel, excel2, sheet, sheet2, save+'.docx')
                
            if self.Rb_excel.isChecked():
                ExcelToExcel_compare(excel, excel2, sheet, sheet2, save+'.xlsx')
                
            QApplication.processEvents()
            time.sleep(1)

            self.textEdit.append("執行完畢!!")
            self.textEdit.moveCursor(QtGui.QTextCursor.End)
            QApplication.processEvents()
            time.sleep(1)

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