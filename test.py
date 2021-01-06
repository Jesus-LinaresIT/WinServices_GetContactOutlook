from win32.lib import win32evtlogutil
from win32 import win32evtlog
from outlook_connector import MSOutlook

import sys
import time

DEBUG = True

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
