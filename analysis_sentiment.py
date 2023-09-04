import logging
import os.path
import shutil
from pandas.core.frame import DataFrame

import win32com.client
import pandas as pd

from openpyxl import Workbook
from openpyxl import load_workbook

import pprint

from wordcloud import WordCloud, STOPWORDS, ImageColorGenerator
import settings
    
'''
EAR - Email AI Reader
Analyse sentiment in an excel file
'''

import logging
import os.path
import shutil
from pandas.core.frame import DataFrame


import pandas as pd
import re               # Remove punctuation
import matplotlib.pyplot as plt

import settings




'''
Do the Actual analysis
'''
def _analyse_client_emails(my_row):
    print(my_row)

    #email_filtered=emails.loc[emails["Parent"]=='ANAIAHHEALTHCARE']
    #print(email_filtered.head())


            #filter main emails and loop

            #Clean up data
                #TODO punctuation
                #TODO remove stop words

            # calc sentiment for this email

            # add to bag of

        # calc sentiment for this folder

        # save sentiment for this folder

def _loop_through_clients():

    # Read data into papers
    emails = pd.read_excel(settings.EMAIL_DATA_DUMP)

    #Group and save our folder names
    email_subjects = emails.groupby("Parent")["Parent"].count()
    #email_subjects.to_excel('.\\output\\folders.xlsx')

    #print(type(email_subjects))
    counter =0

    #Loop over folder names and analyse each one
    for ind in email_subjects:
        print(str(counter)+" : "+str(ind))
        counter+=1

        # client_sentiment = [_analyse_client_emails(x) for x in email_subjects["Parent"]]

        

# simple code to run from command line
if __name__ == '__main__':

    _loop_through_clients()
