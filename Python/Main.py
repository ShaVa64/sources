import numpy as np
import OpenExcel
import sendmessage
import Helpers
from datetime import datetime
import pandas as pd


strFile='HNY_Emails.xlsx'
strSheet='HNY_2021'
data = pd.ExcelFile(strFile)
sheet = data.parse(strSheet)
## get diemnsions of 'sheet"
lines=sheet.shape.__getitem__(0)
cols=sheet.shape.__getitem__(1)
startline=0
endline=lines
num = range(startline, endline)

strBaseSendTime='03/01/2021'
iMinutesStart = 600 # ' 10:00:00'
iMinutesStart = 1140 # ' 19:00:00'
iMinutesStart += 12 
strTime = Helpers.minutes2hour(iMinutesStart)
print('base start time is ' + strTime)
# Init outlook
strSender="shalev@isako.com"
outlook,senderaccount = sendmessage.init_ol(strSender)

for kk in num:
    name = sheet['NAME'][kk]
    mailto = sheet['EMAIL'][kk]    
    strTimeNow = str(datetime.now())
    ## Add 4 min from the start time 
    # +' 06:00 PM'
    iMinutesToSend= iMinutesStart + ((kk+1) * 3)
    strMinutesSendAt =  Helpers.minutes2hour(iMinutesToSend)
    strSendAt= strBaseSendTime + strMinutesSendAt
    #
    sendmessage.send_mail_003(outlook,senderaccount,name,mailto,strSendAt,strTimeNow,kk)

print ('sending ended at ' + str(datetime.now()) + ', last kk is : ' + str(kk))