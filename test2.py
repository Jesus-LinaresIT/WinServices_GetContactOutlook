from outlook_connector import MSOutlook
from settings import Config

import time, sys, schedule, csv, os.path, requests as r


_ifpath_contacts = Config.PATH_CONTACTS
_ifpath_contacts_new = Config.PATH_CONTACTS_NEW

path_files_csv = Config.PATH_CONTACTS_CSV
path_files_csv_new = Config.PATH_CONTACTS_CSV_NEW


DEBUG = 0

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


def getDifferenceLists(contacts):
  #try:
  if not _ifpath_contacts:
    return updateListContact(contacts)
  else:
    with open (path_files_csv, newline ='') as csv_file:
      csvreader = csv.DictReader(csv_file)

      contact_list = []
      for row in csvreader:
        contact_list.append(row)    # get the list of data type OrderedDict from the already created CSV file

    clist = []
    for d in contact_list:
      clist.append(dict(d))     # get a flat dictionary list without being of type OrderedDict

    contact_difference = [item for item in contacts if item not in clist]   #comprehension list for get difference between of two lists

    return saveNewContact(contact_difference, contacts)
  #except:
    #exit()


def updateListContact(contacts):      # Function that only runs once as long as the contacts.csv file is not created
  #try:
  with open (path_files_csv, 'w', newline ='') as new_file:

    header = ['fullName', 'email']
    writeFile = csv.DictWriter(new_file, fieldnames= header)
    writeFile.writeheader()

    for row in contacts:
      writeFile.writerow(row)

  if not _ifpath_contacts_new:
    return sendtoConstantContact()
  #except:
    #exit()


def saveNewContact(contacts_dif, data_contacts):      # Function that only runs once as long as the contacts.csv file is not created
  #try:
  print("ESTOY AQUI")
  with open (path_files_csv_new, 'w', newline ='') as file:

    header = ['fullName', 'email']

    writeFile = csv.DictWriter(file, fieldnames= header)
    writeFile.writeheader()

    for row in contacts_dif:
      writeFile.writerow(row)

    updateListContact(data_contacts)
    sendtoConstantContact()
  #except:
    exit()


def sendtoConstantContact():
  url = Config.URI
  key = Config.API_KEY
  headers = {'Authorization': 'Bearer <{key}>',
              'content-type': 'multipart/form-data'}

  if not _ifpath_contacts_new:
    files = {'file_name': ('contacts.csv', open(path_files_csv, 'rb'))}


    response = r.post(url, headers=headers, files= files)
    print('llegue hasta aqui')
    exit()

  else:
    file = {'file_name': 'Newcontacts.csv',
        'data':('Newcontacts.csv', open(path_files_csv_new, 'rb'))}

    list_ids = [key]

    response = r.post(url, list_ids=list_ids, files= file)
    print('llegue hasta aqui2')
    exit()





contacts = getContacts()
getDifferenceLists(contacts)



#schedule.every(10).seconds.do(saveFile, getSometing)

while 1:
  schedule.run_pending()
  time.sleep(1)