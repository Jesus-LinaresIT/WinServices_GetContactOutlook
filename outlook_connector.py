import win32com.client
import sys
import csv
import time

DEBUG = 0

class MSOutlook:
  def __init__(self):
      self.outlookFound = 0
      try:
        self.oOutlookApp = \
            win32com.client.gencache.EnsureDispatch("Outlook.Application")
        self.outlookFound = 1
      except:
        exit()
        #print("MSOutlook: unable to load Outlook")
      
      self.records = []
      
  def loadContacts(self, keys=None):
      if not self.outlookFound:
          return

      # this should use more try/except blocks or nested blocks
      onMAPI = self.oOutlookApp.GetNamespace("MAPI")
      ofContacts = \
          onMAPI.GetDefaultFolder(win32com.client.constants.olFolderContacts)


      for oc in range(len(ofContacts.Items)):
          contact = ofContacts.Items.Item(oc + 1)
          if contact.Class == win32com.client.constants.olContact:
              if keys is None:
                  # if we were't give a set of keys to use
                  # then build up a list of keys that we will be
                  # able to process
                  # I didn't include fields of type time, though
                  # those could probably be interpreted
                  keys = []
                  for key in contact._prop_map_get_:
                      if isinstance(getattr(contact, key), (int, str, 'ignore')):
                          keys.append(key)
              record = {}
              for key in keys:
                  record[key] = getattr(contact, key)
              if DEBUG:
                  print( oc, record['FullName'])
              self.records.append(record)

