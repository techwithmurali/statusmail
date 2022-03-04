# statusmail

Send Status mails based on data available in Excel on regular basis

Mail sending options - PLSQL or using Linux  - decided by parameter in Globals.py

Excel options to be specified in Globals.py and each excel/mail uniquely identified by Jobname.

In Linux option, html file is generated and is moved to Linux Server.

Separate Job in Linux to read html and to send mail (Refer repository bash for the same) 
