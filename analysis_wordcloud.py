import logging
import os.path
import shutil
from pandas.core.frame import DataFrame

import win32com.client
import pandas as pd

from openpyxl import Workbook
from openpyxl import load_workbook

from wordcloud import WordCloud, STOPWORDS, ImageColorGenerator
import settings
    
'''
EAR - Email AI Reader
Support for managing an outlook inbox - this could be a personal, or share one

Capture information from Outlook into an Excel file
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
Generate a wordcloud
'''
def _do_analysis():

    # Read data into papers
    emails = pd.read_excel(settings.EMAIL_DATA_DUMP)

    # Load the regular expression library
    email_body= ''.join(emails['Body'].to_string())
    print(email_body)

    # Create and generate a word cloud image:
    wordcloud = WordCloud(width=3200, height=1600).generate(email_body)
    wordcloud.to_file(".\\reports\\outputput_full_wordcloud.png")

    # Display the generated image:
    #plt.imshow(wordcloud, interpolation='bilinear')
    #plt.axis("off")
    #plt.show()


# simple code to run from command line
if __name__ == '__main__':

    _do_analysis()
