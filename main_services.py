from base_service import SMWinservice
from win32.lib import win32evtlogutil
from win32 import win32evtlog
from outlook_connector import MSOutlook

import time
import sys
import csv
import schedule

DEBUG = 0

class PythonContactConstant(SMWinservice):
  _svc_name_ = "ContactConstant"
  _svc_display_name_ = "ContactConstant Winservice"
  _svc_description_ = "Get outlook contact"

  def start(self):
    self.isrunning = True

  def stop(self):
    self.isrunning = False

  def main(self):
    contact = getContacts()

    schedule.every(1).minutes.do(saveFile, contact)
    #schedule.every().days().do(saveFile, contact)

    while 1:
      schedule.run_pending()
      time.sleep(1)

def getContacts():
    oOutlook = MSOutlook()
    # delayed check for Outlook on win32 box
    if not oOutlook.outlookFound:
      sys.exit(1)
  
    fields = ['FullName','Email1Address']

    if DEBUG:
      startTime = time.time()

    # you can either get all of the data fields
    # or just a specific set of fields which is much faster
    #oOutlook.loadContacts()
    oOutlook.loadContacts(fields)

    contacts = []
    for contact in oOutlook.records:
      contacts.append({
          "fullName": contact['FullName'],
          "email": contact['Email1Address']
      })

    return contacts

def saveFile(data):
  try:
    with open ("C:\\contacts\\contacts.csv", 'w', newline ='') as new_file:

      header = ['fullName', 'email']
      writeFile = csv.DictWriter(new_file, fieldnames= header)
      writeFile.writeheader()
      print("ESTOY AQUI")

      for row in data:
        print(row)
        writeFile.writerow(row)
  except:
    exit()



if __name__ == '__main__':
  PythonContactConstant.parse_command_line()