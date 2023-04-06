import logging
import os.path
import shutil
from pandas.core.frame import DataFrame

import win32com.client
import pandas as pd

from openpyxl import Workbook
from openpyxl import load_workbook

import settings
import ear_excel
import ear_email
    
'''
EAR - Email AI Reader
Support for managing an outlook inbox - this could be a personal, or share one
'''

# simple code to run from command line
if __name__ == '__main__':
    
    ## Module level variables
    counter=0

    #Handle TO Outlook, Logs and other objects we will need later
    OUTLOOK = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    #Set the Logging level. Change it to logging.INFO is you want just the important info
    logging.basicConfig(filename=settings.LOG_FILE, encoding='utf-8', level=logging.DEBUG)

    #Set the working directory
    os.chdir(settings.WORKING_DIRECTORY)
    print ("\nSet working directory to: "+os.getcwd())

    # Carry out the steps to sync excel adn outlook
    # ear_excel.clear_excel_output_file()
    ear_email.export_email_to_excel(OUTLOOK)
    
