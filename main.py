from outlook_se.servi.base_service import SMWinservice
from outlook_se.servi.contact_services import win_service
from win32.lib import win32evtlogutil
from win32 import win32evtlog


import schedule, time

get_contact = win_service.getContacts()


class PythonContactConstant(SMWinservice):
  _svc_name_ = "ContactConstant"
  _svc_display_name_ = "ContactConstant Winservice"
  _svc_description_ = "Get outlook contact"

  def start(self):
    self.isrunning = True

  def stop(self):
    self.isrunning = False

  def main(self):

    schedule.every(45).seconds.do(win_service.getDifferenceLists, get_contact)
    #schedule.every(1).minutes.do(win_service.getDifferenceLists, get_contact)
    #schedule.every().days().do(win_service.getDifferenceLists, get_contact)

    while 1:
      schedule.run_pending()
      time.sleep(1)


if __name__ == '__main__':
  PythonContactConstant.parse_command_line()