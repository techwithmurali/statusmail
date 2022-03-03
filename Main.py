"""Status Mail Generation Utility.

Based on the parameters specified in the Globals.py file this program generates mail from the 
details present in the excel.ExcelName,Sheets and columns that should be mailed can be configured 
in Globals.py.

Prerequiste for execution:
Globals.g_global_dict to be maintained with a unique key which will be used as jobname.
Invoke genStatusMail with the jobname
Global parameters will be fetched for the input jobname.
Excel and columns will be fetched from globals and mail would get generated.
This is scheduled on regular intervals.

"""

import schedule,os
import time
from datetime import datetime
import Mail_Utils as mail
import Globals

def get_curr_time():
    """ Returns the current time """
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    return dt_string

def genStatusMail(pjobName):
    """ Invoked from Job,this function generates status mail. 

    Parameters:
    jobname:JOB1 
        The Necessary global parameters for this jobname should be available in Globals.py in g_global_dict
        This function extracts necessary parameters for this jobname and generates status mail.
        Following parameters are expected:
        ExcelPath and ExcelName ==> Path and Name of Excel using which status mail needs to be generated.
        Work_dir_path           ==> Working Directory
        SheetList               ==> List of sheets which needs to be considered for mail generation.
        MsgBody                 ==> Message Body
        Subject                 ==> Mail Subject
        Footer_text             ==> Footer Text if any to be added to mail.
        Sender Email            ==> email address of sender.
        l_dict_cols_reqd        ==> Dict with keys as each sheetnames and column numbers in each sheet that needs to be extracted.
        ToList                  ==> recipents email address
        dictSheetDesc           ==> Dict with description of each sheet indexed by sheetnames.

    """

    if Globals.g_global_dict.get(pjobName,'*') == '*':
        print('Global Parameters Not Set for {}'.format(pjobName))
        return 
    l_dict_details = Globals.g_global_dict[pjobName]
    mail.conv_Excel_html(l_dict_details)
    
    ExcelPath   = l_dict_details.get('ExcelPath','')
    ExcelName   = l_dict_details.get('ExcelName','')
    sender      = l_dict_details.get('Sender','')
    Subject     = l_dict_details.get('Subject','')
    daily_status_html   = '{}.html'.format(os.path.splitext(ExcelName)[0])
    subject     = '{}   {} '.format(Subject,get_curr_time())
    ToList      = l_dict_details.get('ToList','') 
    Globals.pr_sendMail_Plsql(sender,subject,os.path.join(ExcelPath,daily_status_html),ExcelPath,ToList)
    print('status sending completed... at {}'.format(get_curr_time()))

def schedule_job():
    """ Executes the sub function job at specific intervals """
    print('\n \n \n \n')
    print('TD Java Status Mails - Scheduled at 4:45 PM: Please check with Murali before closing this...')
    
    def job():
        print('Start of  Status Mail Generation')
        genStatusMail('JOB1')
        genStatusMail('JOB2')
        print('End  of Status Mail Generation')        
        now = datetime.now()
        dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
        print('status sending completed... at {}'.format(dt_string))
        print('\n \n *********************************************** \n \n ')
        
    schedule.every().tuesday.at("16:45").do(job)
    schedule.every().thursday.at("16:45").do(job)
    #job()

    
    while True:
        schedule.run_pending()
        time.sleep(60)
        
if __name__ == "__main__":
    schedule_job()

