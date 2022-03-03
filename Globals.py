import cx_Oracle,os
import sys,traceback

conn_str =  'USERNAME/PASSWORD@INSTANCE'  ## FOR MAIL GENERATION USING PLSQL 

g_global_dict = {}
''' *******************************************************************************************'''
l_dict_details = {}
l_dict_details['Work_dir_path'] = r'C:\Murali\WorkDir'
l_dict_details['ExcelPath'] = 'C:\\Murali\\StatusMailArtifacts'
l_dict_details['ExcelName'] = 'ExcelToMail.xlsx'
l_dict_details['SheetList'] = ['Sheet1','Sheet2']
l_dict_details['MsgBody']   = " \n Please find the latest status of Open Issues: "
l_dict_details['Subject']   = 'Outstanding Issues as on - '
l_dict_details['Footer_text'] = 'Please update the latest status @: \\SharedPath\SharedFolder\ExcelToMail.xlsx'
l_dict_details['Sender']      = 'issues_testing_status@noreply.techwithmurali.com'
l_dict_cols_reqd = {}
l_dict_cols_reqd['Sheet1']           = [1,2,3,4,5,6,7]
l_dict_cols_reqd['Sheet2']             = [3,4,5,6,7,8,9,10,11]
l_dict_details['ColsReqd']              = l_dict_cols_reqd
l_dict_details['ToList']                = 'xxxxxx@xxx.com' 
dictSheetDesc = {}
dictSheetDesc['Sheet1']           = 'Short Description - Sheet1'
dictSheetDesc['Sheet2']             = 'Short Description - Sheet2'
l_dict_details['dictSheetDesc']             = dictSheetDesc

g_global_dict['JOB1'] = l_dict_details
'''***********************************************************************************************'''
l_dict_details = {}
l_dict_details['Work_dir_path'] = r'C:\Murali\TestDir'
l_dict_details['ExcelPath'] = 'C:\\Murali\\TestMailArtifacts'
l_dict_details['ExcelName'] = 'TestExcelToMail.xlsx'
l_dict_details['SheetList'] = ['Sheet1','Sheet2']
l_dict_details['MsgBody']   = " \n Please find the latest status of Open Issues (Testing): "
l_dict_details['Subject']   = 'Outstanding Issues as on - '
l_dict_details['Footer_text'] = 'Please update the latest status @: \\SharedPath\SharedFolder\TestExcelToMail.xlsx'
l_dict_details['Sender']      = 'issues_testing_status@noreply.techwithmurali.com'
l_dict_cols_reqd = {}
l_dict_cols_reqd['Sheet1']           = [1,2,3,4,5,6,7]
l_dict_cols_reqd['Sheet2']             = [3,4,5,6,7,8,9,10,11]
l_dict_details['ColsReqd']              = l_dict_cols_reqd
l_dict_details['ToList']                = 'xxxxxx@xxx.com' 
dictSheetDesc = {}
dictSheetDesc['Sheet1']           = 'Short Description - Sheet1'
dictSheetDesc['Sheet2']             = 'Short Description - Sheet2'
l_dict_details['dictSheetDesc']             = dictSheetDesc

g_global_dict['JOB2'] = l_dict_details



'''***********************************************************************************************'''

def pr_sendMail_Plsql(p_sender,p_subject,p_Status_html,p_attchment_path,To_List):
    try:
        p_recipents = To_List
        fp  = open(p_Status_html)
        p_html_body = fp.read()
        fp.close()
        print('Start of pr_sendMail_Plsql')
        print('p_recipents',p_recipents)
        print('p_subject',p_subject)
        print('p_html_body length ',str(len(p_html_body)))
        p_attchment_path = '*'
        l_conn_str = conn_str
        connection1 = cx_Oracle.connect(l_conn_str)
        cur1 = connection1.cursor()
        cur1.callproc('pkg_Send_mails.pr_send_mails',(p_sender,p_recipents, p_subject,p_html_body, p_attchment_path))
        cur1.close()
        connection1.close()
        print('completed pr_sendMail_Plsql')
    except:
        print('********* Failed in pr_sendMail_Plsql  **********')
        print('Unexpected error : {0}'.format(sys.exc_info()[0]))
        traceback.print_exc()
