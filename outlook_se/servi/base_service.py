import socket
import traceback
import os

from win32.lib import win32serviceutil as service_util

import win32.servicemanager as service_manager
import win32.win32event as win32event
import win32.win32service as win32service


class SMWinservice(service_util.ServiceFramework):
    '''Base class to create winservice in Python'''

    _svc_name_ = 'pythonService'
    _svc_display_name_ = 'Python Service'
    _svc_description_ = 'Python Service Description'

    @classmethod
    def parse_command_line(cls):
        '''
        ClassMethod to parse the command line
        '''
        service_util.HandleCommandLine(cls)

    def __init__(self, args):
        '''
        Constructor of the winservice
        '''
        service_util.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def SvcStop(self):
        '''
        Called when the service is asked to stop
        '''
        self.stop()
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        '''
        Called when the service is asked to start
        '''
        self.start()
        service_manager.LogMsg(service_manager.EVENTLOG_INFORMATION_TYPE,
                            service_manager.PYS_SERVICE_STARTED,
                            (self._svc_name_, ''))
        try: # try main
            self.main()
        except:
            service_manager.LogErrorMsg(traceback.format_exc()) # if error print it to event log
            os._exit(-1)  # return some value other than 0 to os so that service knows to restart


    def start(self):
        '''
        Override to add logic before the start
        eg. running condition
        '''
        pass

    def stop(self):
        '''
        Override to add logic before the stop
        eg. invalidating running condition
        '''
        pass

    def main(self):
        '''
        Main class to be ovverridden to add logic
        '''
        pass

# entry point of the module: copy and paste into the new module
# ensuring you are calling the "parse_command_line" of the new created class
if __name__ == '__main__':
  SMWinservice.parse_command_line()