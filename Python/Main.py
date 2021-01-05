import numpy as np
import OpenExcel
import sendmessage
import Helpers
from datetime import datetime
import pandas as pd


iMaxToSend = 2
strBaseSendTime='05/01/2021'
iMinutesStart = 720 # ' 12:00:00'
iMinutesStart += 4 
# iMinutesStart = 600 # ' 10:00:00'
# iMinutesStart = 1140 # ' 19:00:00'

strFile='HNY_Emails.xlsx'
strSheet='HNY_2021'
strTypesToSend = '_TEST_DIDI_'
strTypes_NOT_ToSend = '_TEST_DIDI_'

data = pd.ExcelFile(strFile)
sheet = data.parse(strSheet)
## get diemnsions of 'sheet"
lines=sheet.shape.__getitem__(0)
cols=sheet.shape.__getitem__(1)
startline=0
endline=lines
num = range(startline, endline)

iSeconds=randrange(0, 1, 1)
strTime = Helpers.minutes2hour(iMinutesStart,iSeconds)
print('base start time is ' + strTime)
# Init outlook
strSender="shalev@isako.com"
outlook,senderaccount = sendmessage.init_ol(strSender)
iToSend=1
for kk in num:
    strInfos = sheet['INFOS'][kk]
    strSalut = sheet['SALUT'][kk]
    strEmails = sheet['EMAILS'][kk]    
    strSingle = sheet['SINGLE/PLU'][kk] 
    isSingle = True if strSingle == 'S' else False
    strVousTu = sheet['VOUS/TU'][kk] 
    isVous = True if strVousTu == 'V' else False
    strAjout = sheet['AJOUT'][kk] 
    strSignature = sheet['SIGNATURE'][kk] 
    strType = sheet['TYPE'][kk] 
    strDontSend= sheet['DONT-SEND'][kk] 
    bDontSend = True if strDontSend.strip() != '' else False
    strSent = sheet['SENT'][kk] 
    bAlreadySent = True if strSent.strip() != '' else False
    strRetour = sheet['RETOUR'][kk] 

    bTypeNotOK = False if strTypes_NOT_ToSend.find('_'+strType+'_')>=0 else True
    bTypeOK = True if strTypesToSend.find('_'+strType+'_')>=0 else False

    if (not bTypeOK) or (bTypeNotOK) or (bDontSend) or (bAlreadySent): 
        print ('line='+str(kk)+',SKIP :'+strInfos+' / ' +strEmails)
        continue
    #
    strTimeNow = str(datetime.now())
    ## Add 4 min from the start time 
    # +' 06:00 PM'
    iMinutesToSend= iMinutesStart + (iToSend * 3)
    strMinutesSendAt =  Helpers.minutes2hour(iMinutesToSend)
    strToSendAt= strBaseSendTime + strMinutesSendAt
    #
    strEmailSubject='Belle et heureuse année 2021 !' 
    if strType == 'TEST':
        strEmailSubject += ' (id=['+ str(kk) +'/'+ str(iToSend)+ '], created=[' + strTimeNow + '],  to send=[' + strToSendAt +  '] )' 
    #
    strHTMLBody = Helpers.format_HTML_body(strSalut,isSingle,isVous,strAjout,strSignature,strType)
    #
    sent = sendmessage.send_mail_003(outlook,senderaccount,strEmails,strEmailSubject,strHTMLBody,strType,strToSendAt,strTimeNow,kk,iMessagesSent)
    # Do not send tot many at once
    if sent :
        iToSend += 1
        iMessagesSent += 1
        if iMessagesSent >=iMaxToSend :
            break
    #

print ('sending ended at ' + str(datetime.now()) + ', last kk is : ' + str(kk)+ ', mesges sent : ' + str(iMessagesSent))