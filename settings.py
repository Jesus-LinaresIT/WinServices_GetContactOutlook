from dotenv import load_dotenv
load_dotenv()

import os

class Config(object):
    URI = os.environ.get('URI')
    API_KEY = os.environ.get('CONSTANT_CONTACT_API_KEY')
    CONSTANT_CONTACT_LIST_IDS = os.environ.get('CONSTANT_CONTACT_LIST_IDS')

    PATH_CONTACTS = os.path.isfile("C:\\contacts\\contacts.csv")
    PATH_CONTACTS_NEW = os.path.isfile("C:\\contacts\\Newcontacts.csv")
    PATH_CONTACTS_CSV = "C:\\contacts\\contacts.csv"
    PATH_CONTACTS_CSV_NEW = "C:\\contacts\\Newcontacts.csv"
