from ssl import ALERT_DESCRIPTION_BAD_CERTIFICATE_STATUS_RESPONSE
from PyQt5.QtGui import * 
from PyQt5.QtWidgets import * 
from PyQt5.QtCore import *
import sys
import time 
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common import keys
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.alert import Alert
import time
import os
import pandas as pd
import datetime
import sqlite3
import pickle
import smtplib
from email.message import EmailMessage
import re

class pickling:
    @staticmethod
    def serialize(Participants_list):
        return pickle.dumps(Participants_list)
    
    @staticmethod
    def unserialize(Participants_blob):
        return pickle.loads(Participants_blob)

class dataBase():
    conn = sqlite3.connect('report.db')

    c = conn.cursor()

    def insertReport(self,report_dict):
        with dataBase.conn:
            dataBase.c.execute("INSERT INTO attendance VALUES (:id, :date, :session, :names, :attendance, :attendees, :absentees,:unknown_participants, :no_of_people_present, :no_of_people_absent)",{'id':report_dict['id'], 'date':report_dict['date'], 'session':report_dict['session'], 'names':pickling.serialize(report_dict['names']),'attendance':pickling.serialize(report_dict['attendance']), 'attendees':pickling.serialize(report_dict['attendees']), 'absentees':pickling.serialize(report_dict['absentees']), 'unknown_participants':pickling.serialize(report_dict['unknown_participants']), 'no_of_people_present':report_dict['no_of_people_present'], 'no_of_people_absent':report_dict['no_of_people_absent']})
            print('sucess')
    def displayAllReports(self):
        with dataBase.conn:
            dataBase.c.execute("SELECT * FROM attendance")
            return dataBase.c.fetchall()

    def displaySingleReport(self,id):
        with dataBase.conn:
            dataBase.c.execute("SELECT * FROM attendance WHERE id=:id",{'id':id})
            return dataBase.c.fetchall()

    def removeReport(self,id):
        with dataBase.conn:
            dataBase.c.execute("DELETE FROM attendance WHERE id=:id",{'id':id})
            
    def getId():
        with dataBase.conn:
            dataBase.c.execute("""SELECT id FROM attendance  """)
            return dataBase.c.fetchall()

    

class dataProcessing(dataBase):
    id = None
    date = None
    sessionName = None
    participantsList = list()
    names = list()
    attendance =list()
    attendees = list()
    absentees = list()
    unknownParticipants = list()
    no_of_present = [0]
    no_of_absent = [0]
    report = dict()
    nameAndMail = list()
    
    df=pd.read_excel(r".\namelist\namelist.xlsx")
    

    def __init__(self,session,participantsList):
        print('processing data')
        with open('sessionId.txt','r+') as f:
            dataProcessing.id = int(f.read())
            # print(id)
            f.seek(0)
            f.truncate()
            f.write(str(dataProcessing.id+1))

        dataProcessing.date = time.strftime("%d-%m-%Y")
        dataProcessing.sessionName = session
        dataProcessing.participantsList = participantsList
        dataProcessing.names = dataProcessing.df["Names"].values

        for one in dataProcessing.names:
            if one in dataProcessing.participantsList:
                dataProcessing.attendance.append('Present')
                dataProcessing.no_of_present[0]+=1
                dataProcessing.attendees.append(one)
            else:
                dataProcessing.attendance.append('Absent')
                dataProcessing.no_of_absent[0]+=1
                dataProcessing.absentees.append(one)

        for one_ in dataProcessing.participantsList:
            if one_ not in dataProcessing.names:
                dataProcessing.unknownParticipants.append(one_)

        print(dataProcessing.participantsList, dataProcessing.attendance, dataProcessing.no_of_present, dataProcessing.attendees, dataProcessing.no_of_absent, dataProcessing.absentees, dataProcessing.unknownParticipants)
        
        dataProcessing.report['id'] = dataProcessing.id
        dataProcessing.report['date'] = dataProcessing.date
        dataProcessing.report['session'] = str(dataProcessing.sessionName)
        dataProcessing.report['names'] = dataProcessing.names
        dataProcessing.report['attendance'] = dataProcessing.attendance
        dataProcessing.report['attendees'] = dataProcessing.attendees
        dataProcessing.report['absentees'] = dataProcessing.absentees
        dataProcessing.report['unknown_participants'] = dataProcessing.unknownParticipants
        dataProcessing.report['no_of_people_present'] = dataProcessing.no_of_present[0]
        dataProcessing.report['no_of_people_absent'] = dataProcessing.no_of_absent[0]
        print(dataProcessing.report)

        self.insertReport(dataProcessing.report)

        if form.checkBox == 2:
            self.sendMail()

    def sendMail(self):

        result = re.split(r"@", form.gmail)
        print(result[0])
        form.gmail = result[0]
        
        nameAndMail = list()
        df=pd.read_excel(r".\namelist\namelist.xlsx")
        Names = df['Names'].values.tolist()
        Email = df['Email'].values.tolist()
        if len(Email)>0:
            print(Names,Email)

            attendance = dataProcessing.attendance

            absentMail = list()
            presentMail = list()

            for (n,e,at) in zip(Names,Email,attendance):
                if at == "Absent":
                    absentMail.append(e)
                else:
                    presentMail.append(e)

            if len(absentMail)>0:
                msg=EmailMessage()
                msg['Subject' ]= 'Reg Attendance'
                msg['From'] = "Don't Reply - Automated Mail"
                msg['To'] = absentMail
                msg.set_content(f'It is to notify that you were absent for the session {str(dataProcessing.sessionName)} on {dataProcessing.date}')

                with smtplib.SMTP_SSL('smtp.gmail.com',465) as e:
                    e.login(form.gmail + '@gmail.com',form.password)
                    e.send_message(msg)
                print('mail sent')

            if len(presentMail)>0:
                msg=EmailMessage()
                msg['Subject' ]= 'Reg Attendance'
                msg['From'] = "Don't Reply - Automated Mail"
                msg['To'] = presentMail
                msg.set_content(f'It is to notify that you were absent for the session {str(dataProcessing.sessionName)} on {dataProcessing.date}')

                with smtplib.SMTP_SSL('smtp.gmail.com',465) as e:
                    e.login(form.gmail + '@gmail.com',form.password)
                    e.send_message(msg)
                print('mail sent')

            self.msg2 = QMessageBox()
            self.msg2.setIcon(QMessageBox.Information)

            self.msg2.setText("Emails have been sent successfully")
            # self.msg2.setInformativeText("Check the exported folder in the same directory")
            self.msg2.setWindowTitle("Mail Sent")
            # self.msg.setDetailedText("The details are as follows:")  
            self.msg2.setStandardButtons(QMessageBox.Ok)
            self.msg2.show()



class form():

    checkBox = None
    gmail = None
    password = None

    def createFormZoom(self):
        self.gmailLabelZoom = QLabel('Gmail Id:', self)
        self.gmailLabelZoom.move(20, 60)
        self.gmailLabelZoom.resize(200, 50)
        self.gmailLabelZoom.setFont(QFont('Arial', 15))


        self.gmailIdZoom = QLineEdit(self)
        self.gmailIdZoom.move(250, 60)
        self.gmailIdZoom.resize(200, 50)
        self.gmailIdZoom.setFont(QFont('Arial', 15))


        self.passwordLabelZoom = QLabel('Password:', self)
        self.passwordLabelZoom.move(20, 120)
        self.passwordLabelZoom.resize(200, 50)
        self.passwordLabelZoom.setFont(QFont('Arial', 15))

        self.passwordZoom = QLineEdit(self)
        self.passwordZoom.setEchoMode(QLineEdit.Password)
        self.passwordZoom.move(250, 120)
        self.passwordZoom.resize(200, 50)
        self.passwordZoom.setFont(QFont('Arial', 15))

        
        self.sessionLabelZoom = QLabel('Session:', self)
        self.sessionLabelZoom.move(20, 180)
        self.sessionLabelZoom.resize(200, 50)
        self.sessionLabelZoom.setFont(QFont('Arial', 15))

        self.session = QLineEdit(self)
        self.session.move(250, 180)
        self.session.resize(200, 50)
        self.session.setFont(QFont('Arial', 15))

        self.meetingLinkLabelZoom = QLabel('Meetling Link:', self)
        self.meetingLinkLabelZoom.move(20, 240)
        self.meetingLinkLabelZoom.resize(200, 50)
        self.meetingLinkLabelZoom.setFont(QFont('Arial', 15))

        self.meetingLinkZoom = QLineEdit(self)
        self.meetingLinkZoom.move(250, 240)
        self.meetingLinkZoom.resize(350, 50)
        self.meetingLinkZoom.setFont(QFont('Arial', 15))


        self.checkBoxMailZoom= QCheckBox(self) 
        self.checkBoxMailZoom.setGeometry(QRect(170, 120, 81, 20))
        self.checkBoxMailZoom.move(20, 310) 
        # self.checkBoxMailZoom.stateChanged.connect(lambda: print('check Zoom',self.checkBoxMailZoom.checkState()))

        self.MailLabelZoom = QLabel('Send Mail', self)
        self.MailLabelZoom.move(50, 296)
        self.MailLabelZoom.resize(200, 50)
        self.MailLabelZoom.setFont(QFont('Arial', 10))


        self.connectButtonZoom = QPushButton('Connect',self)
        self.connectButtonZoom.setGeometry(QRect(200, 350, 150, 45))
        self.connectButtonZoom.clicked.connect(self.connectToZoom)
    
    def connectToZoom(self):
        form.checkBox = self.checkBoxMailZoom.checkState()
        form.gmail = self.gmailIdZoom.text()
        form.password = self.passwordZoom.text()
        
        
        ch_options=Options()
        driver_path='chromedriver.exe'
        ch_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3641.0 Safari/537.36 ")
        ch_options.add_argument("window-size=1024,720")
        ch_options.add_argument("--disable-notifications")

        ch_options.add_experimental_option("prefs", { \
            "profile.default_content_setting_values.media_stream_mic": 1, 
            "profile.default_content_setting_values.media_stream_camera": 1,
            "profile.default_content_setting_values.geolocation": 1, 
            "profile.default_content_setting_values.notifications": 2,
            "download_restrictions": 3
        
    })
        print("Opening the driver")
        driver=webdriver.Chrome(executable_path=driver_path,chrome_options=ch_options)
        print('connecting to Google')
        driver.get(r'https://accounts.google.com/o/oauth2/auth/identifier?client_id=717762328687-iludtf96g1hinl76e4lc1b9a82g457nn.apps.googleusercontent.com&scope=profile%20email&redirect_uri=https%3A%2F%2Fstackauth.com%2Fauth%2Foauth2%2Fgoogle&state=%7B%22sid%22%3A1%2C%22st%22%3A%2259%3A3%3Abbc%2C16%3Ac51291c820784476%2C10%3A1607098594%2C16%3A8181d90b08d0e2c4%2C670c70e18677663e4fc5e355bc806d81b97c1bdd94cc3583cbdd2b3cb02a1310%22%2C%22cdl%22%3Anull%2C%22cid%22%3A%22717762328687-iludtf96g1hinl76e4lc1b9a82g457nn.apps.googleusercontent.com%22%2C%22k%22%3A%22Google%22%2C%22ses%22%3A%22696a343211d44cd79cb26ed80ad4f9e7%22%7D&response_type=code&flowName=GeneralOAuthFlow')
        # https://accounts.google.com/o/oauth2/auth/identifier?client_id=717762328687-iludtf96g1hinl76e4lc1b9a82g457nn.apps.googleusercontent.com&scope=profile%20email&redirect_uri=https%3A%2F%2Fstackauth.com%2Fauth%2Foauth2%2Fgoogle&state=%7B%22sid%22%3A1%2C%22st%22%3A%2259%3A3%3Abbc%2C16%3Ac51291c820784476%2C10%3A1607098594%2C16%3A8181d90b08d0e2c4%2C670c70e18677663e4fc5e355bc806d81b97c1bdd94cc3583cbdd2b3cb02a1310%22%2C%22cdl%22%3Anull%2C%22cid%22%3A%22717762328687-iludtf96g1hinl76e4lc1b9a82g457nn.apps.googleusercontent.com%22%2C%22k%22%3A%22Google%22%2C%22ses%22%3A%22696a343211d44cd79cb26ed80ad4f9e7%22%7D&response_type=code&flowName=GeneralOAuthFlow

        email_id=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//input[@class='whsOnd zHQkBf']")))
        email_id.send_keys(self.gmailIdZoom.text())

        next_=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//button[@class='VfPpkd-LgbsSe VfPpkd-LgbsSe-OWXEXe-k8QpJ VfPpkd-LgbsSe-OWXEXe-dgl2Hf nCP5yc AjY5Oe DuMIQc qIypjc TrZEUc']")))
        next_.click()

        time.sleep(3)

        email_id=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//input[@class='whsOnd zHQkBf']")))
        email_id.send_keys(self.passwordZoom.text())

        next_=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//button[@class='VfPpkd-LgbsSe VfPpkd-LgbsSe-OWXEXe-k8QpJ VfPpkd-LgbsSe-OWXEXe-dgl2Hf nCP5yc AjY5Oe DuMIQc qIypjc TrZEUc']")))
        next_.click()
        time.sleep(1)
        print("signed into google")
        
        print('joining the meeting')
        driver.get(r'https://zoom.us/signin')
        time.sleep(5)
        driver.save_screenshot("2.png")
        google_click=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//a[@class='login-btn login-btn-google']")))
        google_click.click()

        print(1)

        driver.get(self.meetingLinkZoom.text())
        time.sleep(1.5)
        driver.execute_script(f"window.open('{self.meetingLinkZoom.text()}');")
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(2)

        print(2)

        join_browser=driver.find_element_by_xpath(u'//a[text()="Join from Your Browser"]')
        join_browser.click()

        participants_click=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//button[@class='footer-button__button ax-outline']")))
        participants_click.click()
        print('done')
        participants=driver.find_elements_by_class_name("participants-item__display-name")
        participantsList=[participant.text for participant in participants]
        # print(participants_list)
        driver.quit()
        dataProcessing(self.session.text(), participantsList)
        print('data sent')
        driver.quit()

       
            # print('Error')

    
    def createFormGmeet(self):
        self.gmailLabelGmeet = QLabel('Gmail Id:', self)
        self.gmailLabelGmeet.move(20, 60)
        self.gmailLabelGmeet.resize(200, 50)
        self.gmailLabelGmeet.setFont(QFont('Arial', 15))


        self.gmailIdGmeet = QLineEdit(self)
        self.gmailIdGmeet.move(250, 60)
        self.gmailIdGmeet.resize(200, 50)
        self.gmailIdGmeet.setFont(QFont('Arial', 15))


        self.passwordLabelGmeet = QLabel('Password:', self)
        self.passwordLabelGmeet.move(20, 120)
        self.passwordLabelGmeet.resize(200, 50)
        self.passwordLabelGmeet.setFont(QFont('Arial', 15))


        self.passwordGmeet = QLineEdit(self)
        self.passwordGmeet.setEchoMode(QLineEdit.Password)
        self.passwordGmeet.move(250, 120)
        self.passwordGmeet.resize(200, 50)
        self.passwordGmeet.setFont(QFont('Arial', 15))

        self.sessionLabelGmeet = QLabel('Session:', self)
        self.sessionLabelGmeet.move(20, 180)
        self.sessionLabelGmeet.resize(200, 50)
        self.sessionLabelGmeet.setFont(QFont('Arial', 15))

        self.session = QLineEdit(self)
        self.session.move(250, 180)
        self.session.resize(200, 50)
        self.session.setFont(QFont('Arial', 15))

        self.meetingLinkLabelGmeet = QLabel('Meetling Link:', self)
        self.meetingLinkLabelGmeet.move(20, 240)
        self.meetingLinkLabelGmeet.resize(200, 50)
        self.meetingLinkLabelGmeet.setFont(QFont('Arial', 15))

        self.meetingLinkGmeet = QLineEdit(self)
        self.meetingLinkGmeet.move(250, 240)
        self.meetingLinkGmeet.resize(350, 50)
        self.meetingLinkGmeet.setFont(QFont('Arial', 15))


        self.checkBoxMailGmeet= QCheckBox(self) 
        self.checkBoxMailGmeet.setGeometry(QRect(170, 120, 81, 20))
        self.checkBoxMailGmeet.move(20, 310) 
        # self.checkBoxMailGmeet.stateChanged.connect(lambda: print('check Google',self.checkBoxMailGmeet.checkState()))

        self.MailLabelGmeet = QLabel('Send Mail', self)
        self.MailLabelGmeet.move(50, 296)
        self.MailLabelGmeet.resize(200, 50)
        self.MailLabelGmeet.setFont(QFont('Arial', 10))


        self.connectButtonGmeet = QPushButton('Connect',self)
        self.connectButtonGmeet.setGeometry(QRect(200, 350, 150, 45))
        self.connectButtonGmeet.clicked.connect(self.connectToGoogle)
        

    def connectToGoogle(self):
        form.checkBox = self.checkBoxMailGmeet.checkState()
        form.gmail = self.gmailIdGmeet.text()
        form.password = self.passwordGmeet.text()
       
        dataProcessing.sessionName = self.session.text()
        ch_options=Options()
        driver_path='chromedriver.exe'
        ch_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3641.0 Safari/537.36 ")
        ch_options.add_argument("window-size=1280,720")
        ch_options.add_experimental_option("prefs", { \
        "profile.default_content_setting_values.media_stream_mic": 1, 
        "profile.default_content_setting_values.media_stream_camera": 1,
        "profile.default_content_setting_values.geolocation": 1, 
        "profile.default_content_setting_values.notifications": 2
        })
        
        print("Opening the driver")
        driver=webdriver.Chrome(executable_path=driver_path,chrome_options=ch_options)
        print('connecting to Google')
        driver.get(r'https://accounts.google.com/o/oauth2/auth/identifier?client_id=717762328687-iludtf96g1hinl76e4lc1b9a82g457nn.apps.googleusercontent.com&scope=profile%20email&redirect_uri=https%3A%2F%2Fstackauth.com%2Fauth%2Foauth2%2Fgoogle&state=%7B%22sid%22%3A1%2C%22st%22%3A%2259%3A3%3Abbc%2C16%3Ac51291c820784476%2C10%3A1607098594%2C16%3A8181d90b08d0e2c4%2C670c70e18677663e4fc5e355bc806d81b97c1bdd94cc3583cbdd2b3cb02a1310%22%2C%22cdl%22%3Anull%2C%22cid%22%3A%22717762328687-iludtf96g1hinl76e4lc1b9a82g457nn.apps.googleusercontent.com%22%2C%22k%22%3A%22Google%22%2C%22ses%22%3A%22696a343211d44cd79cb26ed80ad4f9e7%22%7D&response_type=code&flowName=GeneralOAuthFlow')
        # https://accounts.google.com/o/oauth2/auth/identifier?client_id=717762328687-iludtf96g1hinl76e4lc1b9a82g457nn.apps.googleusercontent.com&scope=profile%20email&redirect_uri=https%3A%2F%2Fstackauth.com%2Fauth%2Foauth2%2Fgoogle&state=%7B%22sid%22%3A1%2C%22st%22%3A%2259%3A3%3Abbc%2C16%3Ac51291c820784476%2C10%3A1607098594%2C16%3A8181d90b08d0e2c4%2C670c70e18677663e4fc5e355bc806d81b97c1bdd94cc3583cbdd2b3cb02a1310%22%2C%22cdl%22%3Anull%2C%22cid%22%3A%22717762328687-iludtf96g1hinl76e4lc1b9a82g457nn.apps.googleusercontent.com%22%2C%22k%22%3A%22Google%22%2C%22ses%22%3A%22696a343211d44cd79cb26ed80ad4f9e7%22%7D&response_type=code&flowName=GeneralOAuthFlow

        email_id=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//input[@class='whsOnd zHQkBf']")))
        email_id.send_keys(self.gmailIdGmeet.text())

        next_=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//button[@class='VfPpkd-LgbsSe VfPpkd-LgbsSe-OWXEXe-k8QpJ VfPpkd-LgbsSe-OWXEXe-dgl2Hf nCP5yc AjY5Oe DuMIQc qIypjc TrZEUc']")))
        next_.click()

        time.sleep(3)

        email_id=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//input[@class='whsOnd zHQkBf']")))
        email_id.send_keys(self.passwordGmeet.text())

        next_=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//button[@class='VfPpkd-LgbsSe VfPpkd-LgbsSe-OWXEXe-k8QpJ VfPpkd-LgbsSe-OWXEXe-dgl2Hf nCP5yc AjY5Oe DuMIQc qIypjc TrZEUc']")))
        next_.click()
        time.sleep(1)
        print("signed into google")
        
        print('joining the meeting')
        driver.get(r'https://meet.google.com/?hl=en')

        meet_signin=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//a[@class='glue-header__link ']")))
        meet_signin.click()

        link_=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//input[@class='VfPpkd-fmcmS-wGMbrd B5oKfd']")))
        link_.send_keys(self.meetingLinkGmeet.text())

        join=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//button[@class='VfPpkd-LgbsSe VfPpkd-LgbsSe-OWXEXe-dgl2Hf ksBjEc lKxP2d cjtUbb']")))
        join.click()

        mic=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//div[@class='U26fgb JRY2Pb mUbCce kpROve uJNmj HNeRed QmxbVb']")))
        mic.click()

        cam=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//div[@class='U26fgb JRY2Pb mUbCce kpROve uJNmj QmxbVb FTMc0c N2RpBe jY9Dbb']")))
        cam.click()

        time.sleep(1.5)
        join_now=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//div[@class='uArJ5e UQuaGc Y5sE8d uyXBBb xKiqt']")))
        join_now.click()

        join_now=WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,"//div[@class='uArJ5e UQuaGc kCyAyd QU4Gid foXzLb IeuGXd']")))
        join_now.click()
        print('Joined Meeting')
        print('Obtaining Meeting details')
        time.sleep(2)
        participants = driver.find_elements_by_class_name("ZjFb7c")

        participantsList=[participant.text for participant in participants]
        dataProcessing(self.session.text(), participantsList)
        print('sent')
        driver.quit()

    

  

class Menus():

    def createmenubar(self):
        menuBar = self.menuBar()
        
        joinMenu = menuBar.addMenu("&Join")
        zoomAction = QAction("Zoom", self)
        gmeetAction = QAction("Google Meet", self)
        joinMenu.addAction(zoomAction)
        joinMenu.addAction(gmeetAction)

        zoomAction.triggered.connect(self.zoomActionWindow)
        gmeetAction.triggered.connect(self.gmeetActionWindow)
        # Creating menus using a title
        dbMenu = menuBar.addMenu("&Database")
        exportAction = QAction("View", self)
        removeAction = QAction("Remove or Export", self)
        dbMenu.addAction(exportAction)
        dbMenu.addAction(removeAction)

        exportAction.triggered.connect(self.exportActionWindow)
        removeAction.triggered.connect(self.removeActionWindow)


        helpMenu = menuBar.addMenu("&Help")
        aboutAction = QAction("About", self)
        helpMenu.addAction(aboutAction)

        aboutAction.triggered.connect(self.helpActionWindow)


    def helpActionWindow(self):
        self.h = helpWindow()
        self.h.show()
        self.hide()


    def zoomActionWindow(self):
        self.z = zoomWindow()
        self.z.show()
        self.hide()

    def gmeetActionWindow(self):
        self.g = gmeetWindow()
        self.g.show()
        self.hide()

    def exportActionWindow(self):
        self.e = viewWindow()
        self.e.show()
        self.hide()

    def removeActionWindow(self):
        self.r = exportRemoveWindow()
        self.r.show()
        self.hide()


class helpWindow(QMainWindow, Menus, form):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Attendance Automation System - About")
        self.setGeometry(500, 100, 700, 800)
        self.createmenubar()


        self.Author= QLabel('Author: Christ Oliver LLoyd aka ZeroAutumn',self)
        self.Author.resize(500, 100)
        self.Author.move(20,20)
        self.Author.setFont(QFont('Arial', 10))

        self.Developed= QLabel('I did develop this for my Final year project, To know more about it visit \nmy GitHub',self)
        self.Developed.resize(700, 100)
        self.Developed.move(20,70)
        self.Developed.setFont(QFont('Arial', 10))


        self.DevelopedBy= QLabel(self)
        self.DevelopedBy.setOpenExternalLinks(True)
        self.DevelopedBy.setText("<a href=https://github.com/Zero-Autumn>GitHub &#9829;</a>")
        # self.DevelopedBy.setAlignment(Qt.AlignCenter)
        self.DevelopedBy.resize(500, 100)
        self.DevelopedBy.move(20,110)
        self.DevelopedBy.setFont(QFont('Arial', 10))

        
        # self.About.move(0,10)
       
        


class zoomWindow(QMainWindow, Menus, form):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Attendance Automation System - Zoom")
        self.setGeometry(500, 100, 700, 800)
        self.createmenubar()
        self.createFormZoom()

        


class gmeetWindow(QMainWindow, Menus, form):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Attendance Automation System - Googlemeet")
        self.setGeometry(500, 100, 700, 800)
        self.createmenubar()
        self.createFormGmeet()

class viewWindow(QMainWindow, Menus, dataBase):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Attendance Automation System - Export")
        self.setGeometry(500, 100, 700, 800)
        self.createmenubar()

        self.createTable() 

        

        wid = QWidget(self)
        self.setCentralWidget(wid)
        layout = QGridLayout()
        


        # layout = QGridLayout()
        # Add widgets to the layout
        layout.addWidget(self.tableWidget, 0, 0, 1, 2)

        self.idLineedit=QLineEdit()
        self.idLineedit.setPlaceholderText('Enter the id here')

        self.idButton = QPushButton('Open', self)


        layout.addWidget(self.idLineedit, 1, 0)
        layout.addWidget(self.idButton, 1, 1)
        # Set the layout on the application's window
        # self.setLayout(layout)
        wid.setLayout(layout)

        self.idButton.clicked.connect(self.openSingleReport)
    
    def openSingleReport(self):
       

        ids = dataBase.getId()
        print(ids)
        
        userInId = (int(self.idLineedit.text()),)
        print(userInId)
        
        if userInId in ids:
           print(124)  


           self.tableWidget = QTableWidget()
           self.tableWidget.setEditTriggers(QTableWidget.NoEditTriggers)

           fetched_reports = dataBase.displaySingleReport(self,self.idLineedit.text())
           print(len(pickling.unserialize(fetched_reports[0][3])))

        #Row count 
           self.tableWidget.setRowCount(len(pickling.unserialize(fetched_reports[0][3])))  
        #Column count 
           self.tableWidget.setColumnCount(10)
           Headers = ['Id','Date','Session','Names','Attendance','Attendees','Absentees','Unknown Participants','No of people_present','no_of_people_absent']
        
           self.tableWidget.setHorizontalHeaderLabels(Headers) 
           data = list()

           data = dict()
           data = {'id': None, 
           'date':None,
           'session':None,
           'names':None,
           'attendance':None,
           'attendees':None,
           'absentees':None,
           'unknown_participants':None,
           'no_of_people_present':None,
           'no_of_people_absent':None}


           for report in fetched_reports:
               report=enumerate(report)
               for i,e in report:
                   if i == 0:
                      data['id'] = [e]
                   if i == 1:
                      data['date'] = [e]
                   if i == 2:
                      data['session'] = [e]
                   if i == 3:
                      data['names'] = pickling.unserialize(e).tolist()
                   if i == 4:
                      data['attendance'] = pickling.unserialize(e)
                   if i == 5:
                      data['attendees'] = pickling.unserialize(e)
                   if i == 6:
                      data['absentees'] = pickling.unserialize(e)
                   if i == 7:
                      data['unknown_participants'] = pickling.unserialize(e)
                   if i == 8:
                      data['no_of_people_present'] = [e]
                   if i == 9:
                      data['no_of_people_absent'] = [e]
                

               print(data)
        
        
        
           for n, key in enumerate(data.keys()):
            
               for m, item in enumerate(data[key]):
                   newitem = QTableWidgetItem(str(item))
                   self.tableWidget.setItem(m, n, newitem)
                   print((m, n, item))
                    

   
        #Table will fit the screen horizontally 
           self.tableWidget.horizontalHeader().setStretchLastSection(True) 
           self.tableWidget.horizontalHeader().setSectionResizeMode( 
            QHeaderView.Stretch) 



           wid = QWidget(self)
           self.setCentralWidget(wid)
           layout = QGridLayout()
           layout.addWidget(self.tableWidget, 0, 0, 1, 2)
           wid.setLayout(layout)  
        else:
            self.msg = QMessageBox()
            self.msg.setIcon(QMessageBox.Information)

            self.msg.setText("Invalid Id")
            self.msg.setInformativeText("The ID you have entered can't be found")
            self.msg.setWindowTitle("Error")
            # self.msg.setDetailedText("The details are as follows:")  
            self.msg.setStandardButtons(QMessageBox.Ok)
            self.msg.show()
            # self.msg.buttonClicked.connect(lambda: print(1))
            
            
            print('no')

    #Create table 
    def createTable(self): 


        self.tableWidget = QTableWidget()

        fetched_reports = dataBase.displayAllReports(self)

  
        #Row count 
        self.tableWidget.setRowCount(len(fetched_reports))  
        #Column count 
        self.tableWidget.setColumnCount(3)
        #Headers = ['id','date','session','names','attendance','attendees','absentees','unknown_participants','no_of_people_present','no_of_people_absent']
        Headers = ['Id','Date','Session']
        self.tableWidget.setHorizontalHeaderLabels(Headers) 
        print(fetched_reports)

        for i, report in enumerate(fetched_reports):
            report=enumerate(report)
            for j,e in report:
                newitem = QTableWidgetItem(str(e))
                self.tableWidget.setItem(i, j, newitem)
                print((i,j))
         
        #Table will fit the screen horizontally 
        self.tableWidget.horizontalHeader().setStretchLastSection(True) 
        self.tableWidget.horizontalHeader().setSectionResizeMode( 
            QHeaderView.Stretch) 




class exportRemoveWindow(QMainWindow, Menus, dataBase):
    idline = None
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Attendance Automation System - remove")
        self.setGeometry(500, 100, 700, 800)
        self.createmenubar()

        self.createTable()
        self.Exportbtn = QPushButton('Export') 
        self.Removebtn = QPushButton('Remove') 

        

        wid = QWidget(self)
        self.setCentralWidget(wid)
        layout = QGridLayout()
        


        # layout = QGridLayout()
        # Add widgets to the layout
        layout.addWidget(self.tableWidget, 0, 0, 1, 2)

        self.idLineedit=QLineEdit()
        self.idLineedit.setPlaceholderText('Enter the id here')

        

        self.idButton = QPushButton('Open', self)


        layout.addWidget(self.idLineedit, 1, 0)
        layout.addWidget(self.idButton, 1, 1)
        # Set the layout on the application's window
        # self.setLayout(layout)
        wid.setLayout(layout)

        self.idButton.clicked.connect(self.openSingleReport)
    
    def openSingleReport(self):

        exportRemoveWindow.idline = self.idLineedit.text()
       

        ids = dataBase.getId()
        print(ids)
        
        userInId = (int(self.idLineedit.text()),)
        print(userInId)
        
        if userInId in ids:
           print(124)  


           self.tableWidget = QTableWidget()
           self.tableWidget.setEditTriggers(QTableWidget.NoEditTriggers)

           fetched_reports = dataBase.displaySingleReport(self,self.idLineedit.text())
           print(len(pickling.unserialize(fetched_reports[0][3])))

        #Row count 
           self.tableWidget.setRowCount(len(pickling.unserialize(fetched_reports[0][3])))  
        #Column count 
           self.tableWidget.setColumnCount(10)
           Headers = ['Id','Date','Session','Names','Attendance','Attendees','Absentees','Unknown Participants','No of people_present','no_of_people_absent']
        
           self.tableWidget.setHorizontalHeaderLabels(Headers) 
           data = list()

           data = dict()
           data = {'id': None, 
           'date':None,
           'session':None,
           'names':None,
           'attendance':None,
           'attendees':None,
           'absentees':None,
           'unknown_participants':None,
           'no_of_people_present':None,
           'no_of_people_absent':None}


           for report in fetched_reports:
               report=enumerate(report)
               for i,e in report:
                   if i == 0:
                      data['id'] = [e]
                   if i == 1:
                      data['date'] = [e]
                   if i == 2:
                      data['session'] = [e]
                   if i == 3:
                      data['names'] = pickling.unserialize(e).tolist()
                   if i == 4:
                      data['attendance'] = pickling.unserialize(e)
                   if i == 5:
                      data['attendees'] = pickling.unserialize(e)
                   if i == 6:
                      data['absentees'] = pickling.unserialize(e)
                   if i == 7:
                      data['unknown_participants'] = pickling.unserialize(e)
                   if i == 8:
                      data['no_of_people_present'] = [e]
                   if i == 9:
                      data['no_of_people_absent'] = [e]
                

               print(data)
        
        
        
           for n, key in enumerate(data.keys()):
            
               for m, item in enumerate(data[key]):
                   newitem = QTableWidgetItem(str(item))
                   self.tableWidget.setItem(m, n, newitem)
                   print((m, n, item))
                    

   
        #Table will fit the screen horizontally 
           self.tableWidget.horizontalHeader().setStretchLastSection(True) 
           self.tableWidget.horizontalHeader().setSectionResizeMode( 
            QHeaderView.Stretch) 



           wid = QWidget(self)
           self.setCentralWidget(wid)
           layout = QGridLayout()
           layout.addWidget(self.tableWidget, 0, 0, 1, 2)


           
           layout.addWidget(self.Exportbtn, 1, 0)
           layout.addWidget(self.Removebtn, 1, 1)

           self.Exportbtn.clicked.connect(self.export)
           self.Removebtn.clicked.connect(self.remove) 

           wid.setLayout(layout)

           

        else:
            self.msg = QMessageBox()
            self.msg.setIcon(QMessageBox.Information)

            self.msg.setText("Invalid Id")
            self.msg.setInformativeText("The ID you have entered can't be found")
            self.msg.setWindowTitle("Error")
            # self.msg.setDetailedText("The details are as follows:")  
            self.msg.setStandardButtons(QMessageBox.Ok)
            self.msg.show()
            # self.msg.buttonClicked.connect(lambda: print(1))
            
            
            print('no')

    def remove(self):
        usrid =  exportRemoveWindow.idline
        print(usrid)
        self.removeReport(usrid)

        self.msg = QMessageBox()
        self.msg.setIcon(QMessageBox.Information)

        self.msg.setText("Report removed")
        self.msg.setInformativeText("The Report has been removed succesfully")
        self.msg.setWindowTitle("Removed Succesfully")
        # self.msg.setDetailedText("The details are as follows:")  
        self.msg.setStandardButtons(QMessageBox.Ok)
        self.msg.show()

        self.e = viewWindow()
        self.e.show()
        self.hide()


    def export(self):

        usrid =  exportRemoveWindow.idline
        print(usrid)
        fetched_reports = dataBase.displaySingleReport(self,usrid)
        
        data = {'id': None, 
           'date':None,
           'session':None,
           'names':None,
           'attendance':None,
           'attendees':None,
           'absentees':None,
           'unknown_participants':None,
           'no_of_people_present':None,
           'no_of_people_absent':None}


        for report in fetched_reports:
            report=enumerate(report)
            for i,e in report:
                if i == 0:
                    data['id'] = [e]
                if i == 1:
                    data['date'] = [e]
                if i == 2:
                    data['session'] = [e]
                if i == 3:
                    data['names'] = pickling.unserialize(e).tolist()
                if i == 4:
                    data['attendance'] = pickling.unserialize(e)
                if i == 5:
                    data['attendees'] = pickling.unserialize(e)
                if i == 6:
                    data['absentees'] = pickling.unserialize(e)
                if i == 7:
                    data['unknown_participants'] = pickling.unserialize(e)
                if i == 8:
                    data['no_of_people_present'] = [e]
                if i == 9:
                    data['no_of_people_absent'] = [e]
            
            name = '\\' + str(data['id'][0]) +' '+ str(data['date'][0]) +' '+ str(data['session'][0]) + '.xlsx'
            print(data)

            lid,ldate,lsession,lnames,lattendance,lattendees,labsentees,lunknown_participants,lno_of_people_present,lno_of_people_absent = len(data['id']),len(data['date']),len(data['session']),len(data['names']),len(data['attendance']),len(data['attendees']),len(data['absentees']),len(data['unknown_participants']),len(data['no_of_people_present']),len(data['no_of_people_absent'])
            max_len = max(lid,ldate,lsession,lnames,lattendance,lattendees,labsentees,lunknown_participants,lno_of_people_present,lno_of_people_absent)
            if not max_len == lid:
               data['id'].extend(['']*(max_len-lid))
            
            if not max_len == ldate:
               data['date'].extend(['']*(max_len-ldate))

            if not max_len == lsession:
               data['session'].extend(['']*(max_len-lsession))
            
            if not max_len == lnames:
               data['names'].extend(['']*(max_len-lnames))
            
            if not max_len == lattendance:
               data['attendance'].extend(['']*(max_len-lattendance))
            
            if not max_len == lattendees:
               data['attendees'].extend(['']*(max_len-lattendees))
            
            if not max_len == labsentees:
               data['absentees'].extend(['']*(max_len-labsentees))
            
            if not max_len == lunknown_participants:
               data['unknown_participants'].extend(['']*(max_len-lunknown_participants))
            
            if not max_len == lno_of_people_present:
               data['no_of_people_present'].extend(['']*(max_len-lno_of_people_present))
            
            if not max_len == lno_of_people_absent:
               data['no_of_people_absent'].extend(['']*(max_len-lno_of_people_absent))
            
            Final= pd.DataFrame(data)
            path = r".\exported"
            file = path + name
            Final.to_excel(file)

            print('exported')
            self.msg1 = QMessageBox()
            self.msg1.setIcon(QMessageBox.Information)

            self.msg1.setText("The Report has been Exported successfully")
            self.msg1.setInformativeText("Check the exported folder in the same directory")
            self.msg1.setWindowTitle("Export Succesfull")
            # self.msg.setDetailedText("The details are as follows:")  
            self.msg1.setStandardButtons(QMessageBox.Ok)
            self.msg1.show()


       
    #Create table 
    def createTable(self): 

        

        self.tableWidget = QTableWidget()

        fetched_reports = dataBase.displayAllReports(self)

  
        #Row count 
        self.tableWidget.setRowCount(len(fetched_reports))  
        #Column count 
        self.tableWidget.setColumnCount(3)
        #Headers = ['id','date','session','names','attendance','attendees','absentees','unknown_participants','no_of_people_present','no_of_people_absent']
        Headers = ['Id','Date','Session']
        self.tableWidget.setHorizontalHeaderLabels(Headers) 
        print(fetched_reports)

        for i, report in enumerate(fetched_reports):
            report=enumerate(report)
            for j,e in report:
                newitem = QTableWidgetItem(str(e))
                self.tableWidget.setItem(i, j, newitem)
                print((i,j))
                
                

              
  
         
   
        #Table will fit the screen horizontally 
        self.tableWidget.horizontalHeader().setStretchLastSection(True) 
        self.tableWidget.horizontalHeader().setSectionResizeMode( 
            QHeaderView.Stretch)


class mainWindow(QMainWindow, Menus, form):                           
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Attendance Automation System - Google Meet")
        self.setGeometry(500, 100, 700, 800)
        self.createmenubar()
        self.createFormGmeet()

class IntroWindow(QMainWindow): 

    def __init__(self): 
        super().__init__()

  
        # set the title 
        self.setWindowTitle("Attendance Automation System - Zero")
        self.setStyleSheet("border-radius: 10px;") 
        self.setWindowOpacity(0.9)
  
        # setting  the geometry of window 
        # setGeometry(left, top, width, height) 
        self.setGeometry(500, 300, 500, 200)
        self.setWindowFlags(Qt.SplashScreen | Qt.FramelessWindowHint) 

        qtRectangle = self.frameGeometry()
        centerPoint = QDesktopWidget().availableGeometry().center()
        qtRectangle.moveCenter(centerPoint)
        self.move(qtRectangle.topLeft())
  
        # creating a label widget 
        self.Intro = QLabel('''Online Attendance Automation System \n for Zoom and Googlemeet''', self)
        self.Intro.setAlignment(Qt.AlignCenter)
        self.Intro.resize(500, 100)
        self.Intro.move(0,40)
        self.Intro.setFont(QFont('Arial Rounded MT Bold', 14))
        # self.Intro.setStyleSheet("color: yellow;") 
        
        self.Developed= QLabel('Developed By',self)
        self.Developed.resize(500, 100)
        self.Developed.move(130,90)
        self.Developed.setFont(QFont('Arial', 10))


        self.DevelopedBy= QLabel(self)
        self.DevelopedBy.setOpenExternalLinks(True)
        self.DevelopedBy.setText('<a href=https://github.com/Zero-Autumn>Christ Oliver Lloyd</a>')
        # self.DevelopedBy.setAlignment(Qt.AlignCenter)
        self.DevelopedBy.resize(500, 100)
        self.DevelopedBy.move(235,90)
        self.DevelopedBy.setFont(QFont('Arial', 10))

        
        # The QTimer::singleShot is used to call a slot/lambda asynchronously after n ms.
        QTimer.singleShot(4000, self.window2)

        self.main_window()

    def main_window(self):
        self.show()
        

    def window2(self):                                             # <===
        self.w = mainWindow()
        self.w.show()
        self.hide()
        # self.show()
        # show all the widgets   
 

  

# create pyqt5 app 
App = QApplication(sys.argv) 
# create the instance of our Window 
window = IntroWindow()
 
# start the app 
sys.exit(App.exec()) 