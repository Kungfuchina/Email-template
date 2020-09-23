# -*- coding: utf-8 -*-
import win32com.client as win32
import re
import time
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.Qt import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

class Email_template(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUI()
    
    def initUI(self):
        # 设置窗口title和icon
        self.setWindowTitle('Email Template')
        self.setWindowIcon(QIcon('./icon/icon.png'))
        self.resize(400,200)
        settings  =  QSettings("set.ini",QSettings.IniFormat)

        groupname1 = settings.value("groupname1")
        groupname2 = settings.value("groupname2")
        groupname3 = settings.value("groupname3")
        groupname4 = settings.value("groupname4")
        groupname5 = settings.value("groupname5")
        groupname6 = settings.value("groupname6")
        groupname7 = settings.value("groupname7")
        groupname8 = settings.value("groupname8")




        showdate = settings.value("showdate")

        # 添加菜单栏
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('File')
        eidtMenu = menubar.addMenu('Config')
        updateMenu = menubar.addMenu('Update')

        subjectdate = QAction('Show date in email subject', self, checkable=True)
        if showdate == "True":
            subjectdate.setChecked(1)
        elif showdate == "False":
            subjectdate.setChecked(0)
        else:
            pass
        subjectdate.triggered.connect(self.toggleMenu)
        eidtMenu.addAction(subjectdate)



        #  添加控件
        self.grouplabel1 = QLabel(self)
        self.grouplabel1.setText(groupname1)
        self.groupbutton1 = QPushButton('Create Email',self)
        self.groupbutton1.clicked.connect(self.group1)
        self.grouplabel2 = QLabel(self)
        self.grouplabel2.setText(groupname2)
        self.groupbutton2 = QPushButton('Create Email',self)
        self.groupbutton2.clicked.connect(self.group2)
        self.grouplabel3 = QLabel(self)
        self.grouplabel3.setText(groupname3)
        self.groupbutton3 = QPushButton('Create Email',self)
        self.groupbutton3.clicked.connect(self.group3)
        self.grouplabel4 = QLabel(self)
        self.grouplabel4.setText(groupname4)
        self.groupbutton4 = QPushButton('Create Email',self)
        self.groupbutton4.clicked.connect(self.group4)
        self.grouplabel5 = QLabel(self)
        self.grouplabel5.setText(groupname5)
        self.groupbutton5 = QPushButton('Create Email',self)
        self.groupbutton5.clicked.connect(self.group5)
        self.grouplabel6 = QLabel(self)
        self.grouplabel6.setText(groupname6)
        self.groupbutton6 = QPushButton('Create Email',self)
        self.groupbutton6.clicked.connect(self.group6)
        self.grouplabel7 = QLabel(self)
        self.grouplabel7.setText(groupname7)
        self.groupbutton7 = QPushButton('Create Email',self)
        self.groupbutton7.clicked.connect(self.group7)
        self.grouplabel8 = QLabel(self)
        self.grouplabel8.setText(groupname8)
        self.groupbutton8 = QPushButton('Create Email',self)
        self.groupbutton8.clicked.connect(self.group8)



        compoundWidget = QWidget()
        hbox1 = QHBoxLayout()
        hbox1.addWidget(self.grouplabel1)
        hbox1.addWidget(self.groupbutton1)
        hbox2 = QHBoxLayout()
        hbox2.addWidget(self.grouplabel2)
        hbox2.addWidget(self.groupbutton2)
        hbox3 = QHBoxLayout()
        hbox3.addWidget(self.grouplabel3)
        hbox3.addWidget(self.groupbutton3)
        hbox4 = QHBoxLayout()
        hbox4.addWidget(self.grouplabel4)
        hbox4.addWidget(self.groupbutton4)
        hbox5 = QHBoxLayout()
        hbox5.addWidget(self.grouplabel5)
        hbox5.addWidget(self.groupbutton5)
        hbox6 = QHBoxLayout()
        hbox6.addWidget(self.grouplabel6)
        hbox6.addWidget(self.groupbutton6)
        hbox7 = QHBoxLayout()
        hbox7.addWidget(self.grouplabel7)
        hbox7.addWidget(self.groupbutton7)
        hbox8 = QHBoxLayout()
        hbox8.addWidget(self.grouplabel8)
        hbox8.addWidget(self.groupbutton8)

        vbox = QVBoxLayout()
        vbox.addLayout(hbox1)
        vbox.addLayout(hbox2)
        vbox.addLayout(hbox3)
        vbox.addLayout(hbox4)
        vbox.addLayout(hbox5)
        vbox.addLayout(hbox6)
        vbox.addLayout(hbox7)
        vbox.addLayout(hbox8)


        compoundWidget.setLayout(vbox) 
        self.setCentralWidget(compoundWidget)
        self.show()

    def toggleMenu(self, state):
        settings  =  QSettings("set.ini",QSettings.IniFormat)
        if state:
            settings.setValue('showdate', str(bool(1)))
        else:
            settings.setValue('showdate', str(bool(0)))

    def group1(self):
        settings  =  QSettings("set.ini",QSettings.IniFormat)
        showdate = settings.value("showdate")
        groupsubject1 = settings.value("groupsubject1")
        groupSendto1 = settings.value("groupSendto1")
        groupCC1 = settings.value("groupCC1")
        text = open(r"group1.txt", 'r', encoding="utf-8")
        body = text.read()
        if showdate == "True":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject1) +" "+str(date)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto1
            mail.Cc = groupCC1
            mail.Subject = subject
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        elif showdate == "False":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject1)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto1
            mail.Cc = groupCC1
            mail.Subject = subject
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        else:
            pass

    def group2(self):
        settings  =  QSettings("set.ini",QSettings.IniFormat)
        showdate = settings.value("showdate")
        groupsubject2 = settings.value("groupsubject2")
        groupSendto2 = settings.value("groupSendto2")
        groupCC2 = settings.value("groupCC2")
        text = open(r"group2.txt", 'r', encoding="utf-8")
        body = text.read()
        if showdate == "True":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject2) +" "+str(date)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto2
            mail.Cc = groupCC2
            mail.Subject = subject
            
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        elif showdate == "False":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject2)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto2
            mail.Cc = groupCC2
            mail.Subject = subject
            
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        else:
            pass

    def group3(self):
        settings  =  QSettings("set.ini",QSettings.IniFormat)
        showdate = settings.value("showdate")
        groupsubject3 = settings.value("groupsubject3")
        groupSendto3 = settings.value("groupSendto3")
        groupCC3 = settings.value("groupCC3")
        text = open(r"group3.txt", 'r', encoding="utf-8")
        body = text.read()
        if showdate == "True":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject3) +" "+str(date)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto3
            mail.Cc = groupCC3
            mail.Subject = subject
            
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        elif showdate == "False":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject3)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto3
            mail.Cc = groupCC3
            mail.Subject = subject
            
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        else:
            pass

    def group4(self):
        settings  =  QSettings("set.ini",QSettings.IniFormat)
        showdate = settings.value("showdate")
        groupsubject4 = settings.value("groupsubject4")
        groupSendto4 = settings.value("groupSendto4")
        groupCC4 = settings.value("groupCC4")
        text = open(r"group4.txt", 'r', encoding="utf-8")
        body = text.read()
        if showdate == "True":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject4) +" "+str(date)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto4
            mail.Cc = groupCC4
            mail.Subject = subject
            
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        elif showdate == "False":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject4)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto4
            mail.Cc = groupCC4
            mail.Subject = subject
            
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        else:
            pass

    def group5(self):
        settings  =  QSettings("set.ini",QSettings.IniFormat)
        showdate = settings.value("showdate")
        groupsubject5 = settings.value("groupsubject5")
        groupSendto5 = settings.value("groupSendto5")
        groupCC5 = settings.value("groupCC5")
        text = open(r"group5.txt", 'r', encoding="utf-8")
        body = text.read()
        if showdate == "True":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject5) +" "+str(date)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto5
            mail.Cc = groupCC5
            mail.Subject = subject
            
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        elif showdate == "False":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject5)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto5
            mail.Cc = groupCC5
            mail.Subject = subject
            
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        else:
            pass

    def group6(self):
        settings  =  QSettings("set.ini",QSettings.IniFormat)
        showdate = settings.value("showdate")
        groupsubject6 = settings.value("groupsubject6")
        groupSendto6 = settings.value("groupSendto6")
        groupCC6 = settings.value("groupCC6")
        text = open(r"group6.txt", 'r', encoding="utf-8")
        body = text.read()
        if showdate == "True":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject6) +" "+str(date)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto6
            mail.Cc = groupCC6
            mail.Subject = subject
            
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        elif showdate == "False":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject6)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto6
            mail.Cc = groupCC6
            mail.Subject = subject
            
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        else:
            pass

    def group7(self):
        settings  =  QSettings("set.ini",QSettings.IniFormat)
        showdate = settings.value("showdate")
        groupsubject7 = settings.value("groupsubject7")
        groupSendto7 = settings.value("groupSendto7")
        groupCC7 = settings.value("groupCC7")
        text = open(r"group7.txt", 'r', encoding="utf-8")
        body = text.read()
        if showdate == "True":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject7) +" "+str(date)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto7
            mail.Cc = groupCC7
            mail.Subject = subject
            
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        elif showdate == "False":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject7)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto7
            mail.Cc = groupCC7
            mail.Subject = subject
            
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        else:
            pass

    def group8(self):
        settings  =  QSettings("set.ini",QSettings.IniFormat)
        showdate = settings.value("showdate")
        groupsubject8 = settings.value("groupsubject8")
        groupSendto8 = settings.value("groupSendto8")
        groupCC8 = settings.value("groupCC8")
        text = open(r"group8.txt", 'r', encoding="utf-8")
        body = text.read()
        if showdate == "True":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject8) +" "+str(date)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto8
            mail.Cc = groupCC8
            mail.Subject = subject
            
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        elif showdate == "False":
            date = time.strftime('%Y%m%d',time.localtime(time.time()))
            subject = str(groupsubject8)
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.GetInspector
            mail.To = groupSendto8
            mail.Cc = groupCC8
            mail.Subject = subject
            
            bodystart = re.search("<body.*?>", mail.HTMLBody)
            mail.HTMLBody = re.sub(bodystart.group(), bodystart.group(), mail.HTMLBody)
            mail.HTMLBody = body + "\n" + mail.HTMLBody
            mail.Display(False)
        else:
            pass


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Email_template()
    sys.exit(app.exec_())