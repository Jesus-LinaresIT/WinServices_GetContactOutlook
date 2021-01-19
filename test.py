from outlook_se.conectors.outlook_connector import MSOutlook
from settings import Config

import time, sys, schedule, csv, os.path, requests, pickle

DEBUG = 0

class GetContactOutlook():

  def __init__(self):
    self.__path_files_csv = Config.PATH_CONTACTS_CSV
    self.__path_files_csv_new = Config.PATH_CONTACTS_CSV_NEW

    self.__url = Config.URI
    self.__key = Config.API_KEY
    self.__list_ids = Config.CONSTANT_CONTACT_LIST_IDS


  def getContacts(self):
    #try:
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
          "email": contact['Email1Address'],
          "fullName": contact['FullName']
      })

    return contacts
    #except Exception as e:
      #print(e)


  def getDifferenceLists(self, contacts):
    try:
      with open (self.__path_files_csv, newline ='') as csv_file:

        csvreader = csv.DictReader(csv_file)
        contact_list = [row for row in csvreader]
          # get the list of data type OrderedDict
          # from the already created CSV file

      clist = [dict(d) for d in contact_list]
      # get a flat dictionary list without being of type OrderedDict

      contact_difference = [item for item in contacts if item not in clist]
      # comprehension list for get difference between of two lists

      return self.saveNewContact(contact_difference, contacts)
    except:
      return self.updateListContact(contacts)


  def updateListContact(self, contacts):
    """Function that only runs once as long
    as the contacts.csv file is not created"""

    #try:
    with open (self.__path_files_csv, 'w', newline ='') as new_file:

      header = ['email', 'fullName']
      writeFile = csv.DictWriter(new_file, fieldnames= header)
      writeFile.writeheader()

      for row in contacts:
        writeFile.writerow(row)
    try:
      with open (self.__path_files_csv_new, newline ='') as csv_file:
        csvreader = csv.DictReader(csv_file)
    except:
      return self.sendtoConstantContact()


  def saveNewContact(self, contacts_dif, data_contacts):
    """Function that only runs once as long as
    the contacts.csv file is not created """

    #try:
    with open (self.__path_files_csv_new, 'w', newline ='') as file:

      header = ['email', 'fullName']

      writeFile = csv.DictWriter(file, fieldnames= header)
      writeFile.writeheader()

      for row in contacts_dif:
        writeFile.writerow(row)

      self.updateListContact(data_contacts) 
      self.saveDataInB(data_contacts)
      self.sendtoConstantContact()
    #except Exception as e:
      #print(e)


  def sendtoConstantContact(self):
    #try:
      headers = {'Authorization': self.__key,
                  'content-type': 'multipart/form-data',
                  'Accept' : 'multipart/form-data'}

      try:
        with open (self.__path_files_csv_new, newline ='') as csv_file:
          csvreader = csv.DictReader(csv_file)

        files = {'file_name': ('Newcontacts.csv', open(self.__path_files_csv_new, 'rb')),
                  'list_ids' : self.__list_ids}
        response = requests.post(self.__url, headers=headers, files= files)
        print(response)

      except:
        files = {'file': ('contacts.csv', open(self.__path_files_csv, 'rb')),
                  'list_ids' : self.__list_ids}

        response = requests.post(self.__url, headers=headers, files= files)
        print("solo me ejecuto una vez porque ya acabe aqui", response)


  def saveDataInB(self, contact):
    try:
      with open ("C:\\contacts\\info_user", 'wb') as file:
        file.write(contact, b'\xDE\xAD\xBE\xEF')
    except Exception as e:
      print(e)

win_service = GetContactOutlook()
get_contact = win_service.getContacts()


schedule.every(20).seconds.do(win_service.getDifferenceLists, get_contact)

while 1:
  schedule.run_pending()
  time.sleep(1)