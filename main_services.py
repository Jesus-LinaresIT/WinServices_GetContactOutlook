from base_service import SMWinservice
from win32.lib import win32evtlogutil
from win32 import win32evtlog
from outlook_connector import MSOutlook

from test2 import win_service

get_contact = win_service.getContacts()
func_difference = win_service.getDifferenceLists()

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

    schedule.every(1).minutes.do(func_difference, get_contact)
    #schedule.every().days().do(saveFile, contact)

    while 1:
      schedule.run_pending()
      time.sleep(1)


if __name__ == '__main__':
  PythonContactConstant.parse_command_line()