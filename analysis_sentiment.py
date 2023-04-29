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
def _do_analysis():

    # Read data into papers
    emails = pd.read_excel(settings.EMAIL_DATA_DUMP)

    email_subjects = emails.groupby("Parent")["Parent"].count()
    email_subjects.to_excel('.\\output\\folders.xlsx')
    

    for state, frame in emails:
     print(f"First 2 entries for {state!r}")
     print("------------------------")
     print(frame.head(2), end="\n\n")

'''
    email_body= emails['Body']


    dir(email_body)

    # Loop through the first four
    for row_num, row in enumerate(email_body):
        print (row)
        if row_num >4 : break
'''

# simple code to run from command line
if __name__ == '__main__':

    _do_analysis()
