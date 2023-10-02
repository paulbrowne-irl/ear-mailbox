import logging
import traceback
from pandas.core.frame import DataFrame

import pandas as pd

import win32com.client
import os.path

from openpyxl import Workbook
from openpyxl import load_workbook

import settings

counter=0

'''
This script , when pointed at an Outlook Mailbox / folder aims to capture by date
* Email Count - how many total emails sent /recieved
* email complexity - how many lines in an email, tripled if there is an attachment.

'''


'''
Save Dataframe to disk
'''
def _save_summary_info(data_frame):

    logging.info("Saving Dataframe size:"+str(data_frame.size))
    try:
        with pd.ExcelWriter(settings.EMAIL_DATA_DUMP,mode='a',if_sheet_exists="replace") as writer:  
            data_frame.to_excel(writer, sheet_name='Sheet1')
            logging.info("Flushed Cache to disk")
        
            
    except Exception as err:
        
        logging.error ("Error when saving data")
        logging.error(traceback.format_exc())
        logging.error ("\n Was attempting to save")
        logging.error(data_frame.tail(settings.FLUSH_AFTER_X_MAILS))

    return data_frame


'''
Walk folder recursively
'''
def _walk_folder(data_frame,parent_folder,this_folder):
    
    global counter
    
    # Walk and print folders
    for folder in this_folder.Folders:
        logging.info (folder.Name)
        
        #Do recursive call to walk sub folder
        data_frame = _walk_folder(data_frame,parent_folder+"::"+folder.Name,folder)

    #Print folder items
    folderItems = this_folder.Items
 
    for mail in folderItems:

        try:
            #Increment the counter and test if we need to break
            counter+=1

            logging.info("Counter:"+str(counter))
            if(settings.BREAK_AFTER_X_MAILS>0 and counter>settings.BREAK_AFTER_X_MAILS):
                logging.info("Breaking ...")
                return data_frame
            
            #do we need to flush cache to disk?
            if(counter%settings.FLUSH_AFTER_X_MAILS==0):
                data_frame = _save_summary_info(data_frame)

            #Filter on mail items only
            if(mail.Class!=43):
                logging.info("Skipping item type:"+str(mail.Class))

            else:
            
                ## get multiple values


                new_row = pd.DataFrame( {
                        #'Parent':[parent_folder],
                        #'Subject':[""+str(mail.Subject)],
                        #'To':[""+str(mail.To)],
                        #'CC':[""+str(mail.CC)],
                        #'Recipients':[""+str(mail.Recipients)],
                        #'RecievedByName':[""+str(mail.ReceivedByName)],
                        #'ConversationTopic':[""+str(mail.ConversationTopic)],
                        #'ConversationID':[""+str(mail.ConversationID)],
                        #'Sender':[""+str(mail.Sender)],
                        #'SenderName':[""+str(mail.SenderName)],
                        #'SenderEmailAddress':[""+str(mail.SenderEmailAddress)],
                        'email.count':1,
                        'attachments.Count':[""+str(mail.attachments.Count)],
                        'Size':[""+str(mail.Size)],
                        #'ConversationIndex':[""+str(mail.ConversationIndex)],
                        #'EntryID':[""+str(mail.EntryID)],
                        #'Parent':[""+str(mail.Parent)],
                        'CreationTime':[""+str(mail.CreationTime)],
                        'ReceivedTime':[""+str(mail.ReceivedTime)],
                        'LastModificationTime':[""+str(mail.LastModificationTime)],
                        # 'Categories':[""+str(mail.Categories)],
                        'Body':[""+len(str(mail.Body))]
                        })
                
                data_frame= data_frame.append(new_row,ignore_index=True)
        except Exception as e:
            logging.error("error when processing item - will continue")
            logging.error(e)

            
            #HTMLBody
            #RTFBody


    return data_frame
           
        

'''
Output from Outlook Into Excel
'''
def capture_email_count_complexity(OUTLOOK):
    
    
    #debugging
    #root_folder = .Folders.Item(1)
    logging.info("Getting handle to outlook");
    root_folder = OUTLOOK.Folders.Item(settings.INBOX_NAME)

    #Create data frame and save to disk to wipe any previous values
    df = pd.DataFrame()
    df.to_excel(settings.EMAIL_SUMMARY)


    #Walk folders
    logging.info("About to walk folder");
    new_data = _walk_folder(df,"",root_folder)

    #Save the final batch of new data
    _save_summary_info(new_data)

    #Print a sample of the data
    logging.info("complete - sample data")
    logging.info(new_data)



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
    logging.info ("\nSet working directory to: "+os.getcwd())

    # Carry out the steps to sync excel adn outlook
    # ear_excel.clear_excel_output_file()
    capture_email_count_complexity(OUTLOOK)
    





