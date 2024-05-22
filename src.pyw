import os.path
import sys
import time
import pyautogui
import re
import subprocess
import threading

from PyQt5.QtCore import QObject, QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QLabel, QGridLayout, QPushButton, QWidget, QLabel, QSystemTrayIcon, QMenu, QAction, QLineEdit, QMessageBox, QScrollArea, QVBoxLayout
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtGui import QIcon

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import os
import sys
import json

global_stop = False
start_scrapping = False
exiting = False

def parse_loadboard(msg) :
  html = msg
  address1 = ""
  idx = html.find('f7bb0c')
  if idx != -1:
     sub = html[idx:]
     subIdx = sub.find('<p')
     if subIdx != -1:
        sub = sub[subIdx:]
        start = sub.find('>')
        if start != -1:
          sub = sub[start:]
          end = sub.find('<')
          if end != -1:
            address1 = sub[1:end].strip()

  address2 = ""
  idx = html.find('5ed533')
  if idx != -1:
     sub = html[idx:]
     subIdx = sub.find('<p')
     if subIdx != -1:
        sub = sub[subIdx:]
        start = sub.find('>')
        if start != -1:
          sub = sub[start:]
          end = sub.find('<')
          if end != -1:
            address2 = sub[1:end].strip()
  
  vehicles = ""
  idx = html.find('Vehicles')
  if idx != -1:
    sub = html[idx:]
    subIdx = sub.find('<li')
    if subIdx != -1:
        sub = sub[subIdx:]
        start = sub.find('>')
        if start != -1:
          sub = sub[start:]
          end = sub.find('<')
          vehicles = sub[1:end].strip()

  available = ""
  idx = html.find('Available')
  if idx != -1:
    sub = html[idx:]
    subIdx = sub.find('<p')
    if subIdx != -1:
        sub = sub[subIdx:]
        start = sub.find('>')
        if start != -1:
          sub = sub[start:]
          end = sub.find('<')
          if end != -1:
            available = sub[1:end].strip()

  price = ""
  idx = html.find('Price')
  if idx != -1:
    sub = html[idx:]
    subIdx = sub.find('<span')
    if subIdx != -1:
        sub = sub[subIdx:]
        start = sub.find('>')
        if start != -1:
          sub = sub[start:]
          end = sub.find('<')
          if end != -1:
            price = sub[1:end].strip()

  shipper = ""
  idx = html.find('Shipper')
  if idx != -1:
    sub = html[idx:]
    subIdx = sub.find('<p')
    if subIdx != -1:
        sub = sub[subIdx:]
        start = sub.find('>')
        if start != -1:
          sub = sub[start:]
          end = sub.find('<')
          if end != -1:
            shipper = sub[1:end].strip()

  link = ""
  idx = html.find('View Details')
  if idx == -1:
     idx = html.find('View Load')
  if idx != -1:
    sub = html[0:idx]
    subIdx = reverse_tail_search(sub, 'href')
    if subIdx != -1:
        sub = sub[subIdx:]
        start = sub.find('"')
        if start != -1:
          sub = sub[start + 1:]
          end = sub.find(' ')
          if end != -1:
            link = sub[0:end].strip()
  return address1, address2, vehicles, available, price, shipper, link


def parse_centraiddispatch(msg) :
  html = msg
  address1 = ""
  idx = html.find('Pickup')
  if idx != -1:
    sub = html[idx:]
    start = sub.find('>')
    if start != -1:
      sub = sub[start:]
      end = sub.find('<')
      if end != -1:
        address1 = sub[1:end].strip()

  address2 = ""
  idx = html.find('Delivery')
  if idx != -1:
    sub = html[idx:]
    start = sub.find('>')
    if start != -1:
      sub = sub[start:]
      end = sub.find('<')
      if end != -1:
        address2 = sub[1:end].strip()

  vehicles = ""
  idx = html.find('Vehicles')
  if idx != -1:
    sub = html[idx:]
    start = sub.find('&nbsp;')
    if start != -1:
      sub = sub[start + 5:]
      end = sub.find('<')
      if end != -1:
        vehicles = sub[1:end].strip()

  available = ""
  idx = html.find('Available')
  if idx != -1:
    sub = html[idx:]
    start = sub.find('>')
    if start != -1:
      sub = sub[start:]
      end = sub.find('<')
      if end != -1:
        available = sub[1:end].strip()

  price = ""
  idx = html.find('Price')
  if idx != -1:
    sub = html[idx:]
    start = sub.find('>')
    if start != -1:
      sub = sub[start:]
      end = sub.find('<')
      if end != -1:
        price = sub[1:end].strip()

    
  shipper = ""
  idx = html.find('Posted By')
  if idx != -1:
    sub = html[idx:]
    start = sub.find('>')
    if start != -1:
      sub = sub[start:]
      end = sub.find('<')
      if end != -1:
        shipper = sub[1:end].strip()

  link = ""
  idx = html.find('View Listing')
  if idx != -1:
    sub = html[0:idx]
    subIdx = reverse_tail_search(sub, 'href')
    if subIdx != -1:
        sub = sub[subIdx:]
        start = sub.find('"')
        if start != -1:
          sub = sub[start + 1:]
          end = sub.find('"')
          if end != -1:
            link = sub[0:end].strip()
  return address1, address2, vehicles, available, price, shipper, link

class Gmail:
    def __init__(self):
        super().__init__()
        self.active = False
        self.msg = ""
        self.id = ""
        self.subject = ""
        self.snippet = ""
        self.sender = ""
        self.body = ""
        
g_Mails = []
mail_active = False

new_index = 0

terminate = False
model_show = False

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

def fetch_emails():
  """fetch emails
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  from_ = ""
  with open("from", "r") as fp:
    from_ = fp.readline().strip()
    fp.close()
  
  from_1 = ""
  with open("from1", "r") as fp:
    from_1 = fp.readline().strip()
    fp.close()
  
  to_ = ""
  with open("to", "r") as fp:
    to_ = fp.readline().strip()
    fp.close()
  
  executable_path = 'scrapper.exe'

  query = from_
  if from_1 != "":
      query = query + " OR from:" + from_1
  print(query)
  args = [to_, query]

  messages = []
  mails = []
  try:
      startupinfo = subprocess.STARTUPINFO()
      startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
      startupinfo.wShowWindow = subprocess.SW_HIDE
      with open(os.devnull, 'w') as devnull:
        # print(args)
        subprocess.run([executable_path] + args, check=True, startupinfo=startupinfo, stdout=devnull, stderr=devnull)
      try:
        with open("pipe.json", "r", encoding='utf-8') as fp:
        # with open("copy-pipe.json", "r", encoding='utf-8') as fp:
          messages = json.load(fp)
          # print(messages)
          fp.close()
      except FileNotFoundError:
          print("FileNotFoundError")
      except json.JSONDecodeError:
          print("JSONDecodeError")
      except Exception as e:
         print("")
  except subprocess.CalledProcessError as e:
      print(f"Executable failed with exit code {e.returncode}")
  if messages is None:
     messages = []
  for email in messages:
      m = Gmail()
      m.id = email['Id']
      m.msg = ''
      m.subject = email['Subject']
      m.snippet = ''
      m.sender = email['Sender']
      m.body = email['Body']
      mails.append(m)
  
  return mails

class Worker(QObject):
    finished = pyqtSignal()

    def __init__(self):
        super().__init__()

    def do_work(self):
        global mail_active
        global g_Mails
        global hook_sent
        global global_stop
        while True:
            if global_stop == True:
              time.sleep(5)
              continue  
            if start_scrapping == False:
              time.sleep(5)
              continue 
            if terminate:
              break
            mails = fetch_emails()
            if global_stop == True:
              time.sleep(5)
              continue            
            if len(mails) == 0:
              time.sleep(5)
              continue
            mail = mails[0]
            last_id = ""          
            try:
              with open("latestId", "r") as fp:
                last_id = fp.readline().strip()
                fp.close()
            except FileNotFoundError:
              last_id = ""

            if last_id == mail.id:
              time.sleep(5)
              continue

            with open("latestId", "w") as fp:
              fp.write(mail.id)
              fp.close()
            g_Mails = []

            global new_index
            new_index = 0
            idx = 0
            if last_id == "":
               idx = len(mails)
            for e in mails:
              if e.id == last_id:
                new_index = idx
              g_Mails.append(e)
              idx = idx + 1
              if e.sender.find('loadboard@superdispatch.com') != -1:
                if e.body.find('View Load') != -1:
                    cp = Gmail()
                    cp.id = e.id
                    cp.msg = ''
                    cp.subject = e.subject
                    cp.snippet = ''
                    cp.sender = e.sender
                    cp.body = e.body[e.body.find('View Load') + 10:]
                    g_Mails.append(cp)
                    idx = idx + 1
            if len(g_Mails) == 0:
               time.sleep(5)
               continue
            mail_active = True
            self.finished.emit()
def on_close():
    global model_show
    global mail_active
    model_show = False
    mail_active = False

class SettingView(QWidget):
    def __init__(self):
        super().__init__()
        screen_width, screen_height = pyautogui.size()
        self.setGeometry(int((screen_width - 320) / 2), int((screen_height - 255) / 2), 320, 255)
        self.setFixedSize(320, 255)
        self.setWindowTitle("Setting")
        from_ = ""
        try:
          with open("from", "r") as fp:
            from_ = fp.readline().strip()
            fp.close()
        except FileNotFoundError:
          from_ = ""

        from_1 = ""
        try:
          with open("from1", "r") as fp:
            from_1 = fp.readline().strip()
            fp.close()
        except FileNotFoundError:
          from_1 = ""

        to_ = ""
        try:
          with open("to", "r") as fp:
            to_ = fp.readline().strip()
            fp.close()
        except FileNotFoundError:
          to_ = ""

        _ = QLabel(self)
        _.setText("Your email: ")
        _.setGeometry(10, 30, 200, 10)
        self.to_ = QLineEdit(self)
        self.to_.setText(to_)
        self.to_.setGeometry(90, 20, 210, 30)

        _ = QLabel(self)
        _.setText("Sender email: ")
        _.setGeometry(10, 70, 200, 10)        
        self.from_ = QLineEdit(self)
        self.from_.setText(from_)
        self.from_.setGeometry(90, 60, 210, 30)

        _ = QLabel(self)
        _.setText("optional email: ")
        _.setGeometry(10, 110, 200, 10)        
        self.from_1 = QLineEdit(self)
        self.from_1.setText(from_1)
        self.from_1.setGeometry(90, 100, 210, 30)

        # _ = QLabel(self)
        # _.setText("Username: ")
        # _.setGeometry(10, 130, 200, 10)
        # self.username = QLineEdit(self)
        # self.username.setText(getpass.getuser())
        # self.username.setGeometry(90, 120, 210, 30)
        # self.username.setReadOnly(True)

        # _ = QLabel(self)
        # _.setText("Password: ")
        # _.setGeometry(10, 170, 200, 10)
        # self.password = QLineEdit(self)
        # self.password.setText("")
        # self.password.setGeometry(90, 160, 210, 30)
        # self.password.setEchoMode(QLineEdit.Password)

        _ = QPushButton(self)
        _.setText("Start")
        _.setGeometry(90, 215, 80, 25)
        _.clicked.connect(self.on_start_button_clicked)

        _ = QPushButton(self)
        _.setText("Close")
        _.setGeometry(180, 215, 80, 25)
        _.clicked.connect(self.on_close_button_clicked)

    def on_close_button_clicked(self):
      global exiting
      exiting = True
      sys.exit(0)
    def on_start_button_clicked(self):
        global start_scrapping
        
        to_ = self.to_.text()
        from_ = self.from_.text()
        from_1 = self.from_1.text()
        
        if to_ == "":
          msgBox = QMessageBox()
          msgBox.setIcon(QMessageBox.Information)
          msgBox.setText("Please fill your email address")
          msgBox.setWindowTitle("Information")
          msgBox.exec()
          return
        if from_ == "":
          msgBox = QMessageBox()
          msgBox.setIcon(QMessageBox.Information)
          msgBox.setText("Please fill sender email address")
          msgBox.setWindowTitle("Information")
          msgBox.exec()
          return
        # if password == "":
        #   msgBox = QMessageBox()
        #   msgBox.setIcon(QMessageBox.Information)
        #   msgBox.setText("Please enter password")
        #   msgBox.setWindowTitle("Information")
        #   msgBox.exec()
        #   return
        with open("from", "w") as fp:
          fp.write(from_)
          fp.close()
        with open("from1", "w") as fp:
          fp.write(from_1)
          fp.close()

        with open("to", "w") as fp:
          fp.write(to_)
          fp.close()
        # if verify_success(getpass.getuser(), password) == False:
        #   msgBox = QMessageBox()
        #   msgBox.setIcon(QMessageBox.Information)
        #   msgBox.setText("Incorrect password")
        #   msgBox.setWindowTitle("Warning")
        #   msgBox.exec()
        #   return
        start_scrapping = True          
        self.hide()

import simpleaudio as sa
def play_wav(file_path):
    wave_obj = sa.WaveObject.from_wave_file(file_path)
    play_obj = wave_obj.play()
    play_obj.wait_done()

class GMailView(QWidget):
    def __init__(self):
        super().__init__()
        screen_width, screen_height = pyautogui.size()
        self.setGeometry(int((screen_width - 960) / 2), int((screen_height - 840) / 2), 960, 840)

        self.setWindowTitle("New Mails")

        layout = QVBoxLayout(self)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("QScrollArea { border: 1px solid gray; }")
        self.scroll_content = QWidget()
        self.scroll_content_layout = QGridLayout(self.scroll_content)
        self.scroll_area.setWidget(self.scroll_content)
        layout.addWidget(self.scroll_area)
    
        self.worker_thread = QThread()
        self.worker = Worker()
        self.worker.moveToThread(self.worker_thread)
        self.worker.finished.connect(self.on_task_finished)

        self.worker_thread.started.connect(self.worker.do_work)
        self.worker_thread.start()

        
    def closeEvent(self, event):
        event.ignore()
        self.hide()
        on_close()
    def wheelEvent(self, event):
        event.accept() 

    def on_task_finished(self):
      self.setMail()  

    def setMail(self):
      global g_Mails
      
      if len(g_Mails) > 0:
        while self.scroll_content_layout.count():
          item = self.scroll_content_layout.takeAt(0)
          widget = item.widget()
          if widget is not None:
              widget.deleteLater()

        n = 0
        for m in g_Mails:
          row = int(n / 3)
          col = n % 3
          w = QWidget(self)
          w.setMinimumSize(320, 300)
          w.setMaximumSize(320, 300)
          color = "#c8f4ef"
          color = "white"
          if n != 0:
             color = "white"
          border_color = "#dcdcdc"
          border_size = 2

          if n < new_index:
             border_color = "red"
             border_size = 3
          w.setStyleSheet("background-color: " + color + "; border: " + str(border_size) + "px solid " + border_color + "; border-radius: 5px;")

          str_sender = m.sender.split("<")[1].split(">")[0].strip()
          sender = QLabel(w)
          sender.setText("from: " + str_sender)
          sender.setGeometry(10, 15, 280, 10)
          sender.setStyleSheet("border: none; font-size: 11px; font-weight: bold; color: #4f4e4f;")

          title = QLabel(w)
          title.setText(m.subject)
          title.setGeometry(10, 27, 280, 40)
          title.setStyleSheet("border: none; font-size: 14px; font-weight: bold; color: grey;")

          with open("log.txt", "a") as file:
              file.write("-" * 100 + "\n")
              file.write(str_sender + "\n")
              file.write("*" * 100 + "\n")
              file.write(m.body + "\n")
              file.write("=" * 100 + "\n")

          (address1, address2, vehicles, available, price, shipper, link) = ("", "", "", "", "", "", "")
          if str_sender == "loadboard@superdispatch.com":
            (address1, address2, vehicles, available, price, shipper, link) = parse_loadboard(m.body)
          if str_sender == "do-not-reply@centraldispatch.com":
            (address1, address2, vehicles, available, price, shipper, link) = parse_centraiddispatch(m.body)

          subIdx = address1.find('(')
          if subIdx != -1:
             address1 = address1[0:subIdx]
          subIdx = address2.find('(')
          if subIdx != -1:
             address2 = address2[0:subIdx]
             address2 = address2.replace("\n", "")
             address2 = address2.replace("\r", "")
          if len(shipper) >= 24:
             shipper = shipper[0:24] + "..."
          
          # _ = QLabel(w)
          # _.setText("Origin: ")
          # _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: grey;")
          # _.setGeometry(20, 70, 280, 25)
          # _ = QLabel(w)
          # _.setText(address1)
          # _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: #04b408;")
          # _.setGeometry(65, 70, 230, 25)

          # _ = QLabel(w)
          # _.setText("Destination: ")
          # _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: grey;")
          # _.setGeometry(20, 95, 280, 25)
          # _ = QLabel(w)
          # _.setText(address2)
          # _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: #04b408;")
          # _.setGeometry(100, 95, 210, 25)

          # _ = QLabel(w)
          # _.setText("Vehicles: ")
          # _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: grey;")
          # _.setGeometry(20, 120, 280, 25)
          # _ = QLabel(w)
          # _.setText(vehicles)
          # _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: black;")
          # _.setGeometry(80, 120, 210, 25)

          # _ = QLabel(w)
          # _.setText("Available: ")
          # _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: grey;")
          # _.setGeometry(20, 145, 280, 25)
          # _ = QLabel(w)
          # _.setText(available)
          # _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: black;")
          # _.setGeometry(90, 145, 210, 25)

          # _ = QLabel(w)
          # _.setText("Price: ")
          # _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: grey;")
          # _.setGeometry(20, 170, 280, 25)
          # _ = QLabel(w)
          # _.setText(price)
          # _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: black;")
          # _.setGeometry(60, 170, 230, 25)

          # _ = QLabel(w)
          # _.setText("Shipper: ")
          # _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: grey;")
          # _.setGeometry(20, 195, 280, 25)
          # _ = QLabel(w)
          # _.setText(shipper)
          # _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: black;")
          # _.setGeometry(80, 195, 230, 25)

          _ = QLabel(w)
          _.setText(address1)
          _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: #04b408;")
          _.setGeometry(20, 70, 280, 25)

          _ = QLabel(w)
          _.setText(address2)
          _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: #04b408;")
          _.setGeometry(20, 95, 280, 25)

          _ = QLabel(w)
          _.setText(vehicles)
          _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: black;")
          _.setGeometry(20, 120, 280, 25)

          _ = QLabel(w)
          _.setText(available)
          _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: black;")
          _.setGeometry(20, 145, 280, 25)

          _ = QLabel(w)
          _.setText(price)
          _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: black;")
          _.setGeometry(20, 170, 280, 25)

          _ = QLabel(w)
          _.setText(shipper)
          _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: black;")
          _.setGeometry(20, 195, 280, 25)

          if link != "":
            _ = QLabel(w)
            _.setText('<a href="' + link + '">View Details</a>')
            _.setGeometry(20, 230, 180, 25)
            _.setStyleSheet("border: none; font-size: 12px; font-weight: bold; color: blue;")
            _.setOpenExternalLinks(True) 

          phone_numbers = re.findall(r'(\d{3}-\d{3}-\d{4})', m.body)
          if len(phone_numbers) == 0:
            phone_numbers = re.findall(r'(\(\d{3}\)\s*\d{3}-\d{4})', m.body)
          if len(phone_numbers) == 0:
            phone_numbers = re.findall(r'\d{10}', m.body)

          # print("Phone number:", phone_numbers)

          if len(phone_numbers):
            phone = QWebEngineView(w)
            phone.setStyleSheet("background-color: " + color + ";")
            html = ""
            for number in phone_numbers:
               if number == '816-974-7002' or str_sender == "loadboard@superdispatch.com":
                  continue
               ele = '<a href="tel:' + number + '" style="background-color: ' + color + '; text-align: center; border: none; font-size: 14px; font-weight: bold; color: blue;">' + number + '</a><br>'
               html = html + ele
            phone.setHtml(html)
            phone.setGeometry(10, 255, 180, 35)

          self.scroll_content_layout.addWidget(w, row, col)
          n = n + 1
      
        col = 3
        if n > 6:
           n = 6
        if n < 3:
           col = n
        row = int(n / 3)
        if (n % 3) != 0:
           row = row + 1
        h = 350 * row
        w = 350 * col
        screen_width, screen_height = pyautogui.size()
        self.setGeometry(self.geometry().x(), self.geometry().y(), w, h)
        self.show()
        self.showNormal()
        self.activateWindow()
        self.raise_()
        thread = threading.Thread(target=play_wav, args=("alarm.wav",))
        thread.start()
        # play_wav()

def verify_success(username, password):
    from win32security import LogonUser
    from win32con import LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT
    try:
        LogonUser(username, None, password, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT)
    except:
        return False
    return True

def action1_triggered():
  global exiting
  exiting = True
  sys.exit(0)

def reverse_tail_search(main_string, substring):
    reversed_main = main_string[::-1]
    reversed_sub = substring[::-1]
    index = reversed_main.find(reversed_sub)
    if index != -1:
        return len(main_string) - index - len(substring)
    else:
        return -1
    
global action1
global toggle

def toggle_triggered():
  global global_stop
  if global_stop:
      global_stop = False
      toggle.setText("Stop")
      tray_icon.setIcon(QIcon("icon.png"))
  else:
      global_stop = True
      toggle.setText("Start")
      tray_icon.setIcon(QIcon("inactive.png"))

if __name__ == "__main__": 
  if os.path.exists("pipe.json"):
    os.remove("pipe.json")
  if os.path.exists("latestId"):
    os.remove("latestId")
  app = QApplication(sys.argv)  
  tray_icon = QSystemTrayIcon()    
  tray_icon.setToolTip('Email Notification')
  tray_menu = QMenu()  
  toggle = QAction("Stop", tray_menu)
  toggle.triggered.connect(toggle_triggered)
  action1 = QAction("Quit", tray_menu)  
  action1.triggered.connect(action1_triggered)
  tray_icon.setIcon(QIcon("icon.png"))
  tray_menu.addAction(toggle)
  tray_menu.addAction(action1)

  tray_icon.setContextMenu(tray_menu)
  tray_icon.show()  
  setting = SettingView()
  setting.show()
  viewer = GMailView()
  sys.exit(app.exec_())
