import numpy as np
import datetime
import random
import OpenExcel
import sendmessage
import Helpers
from datetime import datetime
import pandas as pd


print('''

==========================================''')

iMaxToSend = 20

dtNowAtStart=datetime.now()
#. . . . . . . . . 
# # make sure the right line is the last of the following 2
strBaseSendDate = dtNowAtStart.strftime('%d/%m/%Y')
strBaseSendTime='08/01/2021' ## If a predetermined date
#. . . . . . . . . 
# make sure the right line is the last of the following 2
iMinutesStart = 7 * 60 # If a predetermined time, here rounded to an hour
iMinutesStart =int(dtNowAtStart.strftime('%H'))*60 + int(dtNowAtStart.strftime('%M')) ## "%H:%M:%S.%fZ"
iMinutesStart += 1 # In any case not before 'x'' minutes from now
#. . . . . . . . . 
# 
strBaseSendDate =str(dtNowAtStart.strftime('%H:%M'))
iMinutesBetweenEmail = 3 
#
strFile='HNY_Emails.xlsx'
strSheet='HNY_2021'
strBaseEmailSubject='Belle et heureuse année 2021 !' 
iShowDelayInSecs = 3
strTypesToSend = '__,_xTEST_' # The '__' means it's OK to send where the TYPE Column is empty, 
                             #  The ',' is not necessary, just to make it clearer
                             #  Any valid value should have an underscore before and an uderscore after.
strTypes_NOT_ToSend = '_DADA_DODO_'

data = pd.ExcelFile(strFile)
sheet = data.parse(strSheet)
## get diemnsions of 'sheet"
lines=sheet.shape.__getitem__(0)
cols=sheet.shape.__getitem__(1)
startline=0
endline=lines
num = range(startline, endline)

iSeconds=random.randrange(0, 59, 1)
strTime = Helpers.minutes2hour(iMinutesStart,iSeconds)

print('base start date-time is ' + strBaseSendTime + ' ' + strTime + ', time now at start of round : ' + str(dtNowAtStart.strftime('%d/%m/%Y %H:%M')) + ', (' + strBaseSendDate +')')
print('Excel has '+ str(lines+1) + ' lines.)')
# Init outlook
strSender="shalev@isako.com"
outlook,senderaccount = sendmessage.init_ol(strSender)
iToSend=1
iMessagesSent=0

for kk in num:
    # use the pd.isnull since empty EXCEL cell beacomes 'nan' when translated to string.
    strInfos = '' if pd.isnull(sheet['INFOS'][kk]) else str(sheet['INFOS'][kk]).strip() 
    strSalut = '' if pd.isnull(sheet['SALUT'][kk]) else str(sheet['SALUT'][kk]).strip()
    strEmails = '' if pd.isnull(sheet['EMAILS'][kk]) else str(sheet['EMAILS'][kk]).strip() 
    if strInfos == '':
        strInfos =strSalut + ' / ' + strEmails
    strSingle = '' if pd.isnull(sheet['SINGLE/PLU'][kk]) else str(sheet['SINGLE/PLU'][kk]).strip() 
    isSingle = True if strSingle == 'S' else False
    strVousTu = '' if pd.isnull(sheet['VOUS/TU'][kk]) else str(sheet['VOUS/TU'][kk]).strip() 
    isVous = True if strVousTu == 'V' else False
    strSender = '' if pd.isnull(sheet['SENDER'][kk]) else str(sheet['SENDER'][kk]).strip() 
    strAjout = '' if pd.isnull(sheet['AJOUT'][kk]) else str(sheet['AJOUT'][kk]).strip() 
    strSignature = '' if pd.isnull(sheet['SIGNATURE'][kk]) else str(sheet['SIGNATURE'][kk]).strip() 
    strType = '' if pd.isnull(sheet['TYPE'][kk]) else str(sheet['TYPE'][kk]).strip() 
    strDontSend= '' if pd.isnull(sheet['DONT-SEND'][kk]) else str(sheet['DONT-SEND'][kk]).strip() 
    bDontSend = False if strDontSend == '' else True
    strLastSent = '' if pd.isnull(sheet['LAST SENT'][kk]) else str(sheet['LAST SENT'][kk]).strip() 
    bAlreadySent = False  if strLastSent.strip() == '' else True
    strRetour = '' if pd.isnull(sheet['RETOUR'][kk]) else str(sheet['RETOUR'][kk]).strip() 

    bTypeNotOK = True if strTypes_NOT_ToSend.find('_'+strType+'_')>=0 else False
    bTypeOK = True if strTypesToSend.find('_'+strType+'_')>=0 else False

    if (not bTypeOK) or (bTypeNotOK) or (bDontSend) or (bAlreadySent): 
        print ('line='+str(kk)+', ----> skip : ' + strInfos + ' / ' + strEmails + '[TypeOK='+ str(bTypeOK)+ ' /TypeNotOK: ' +str(bTypeNotOK)+ ' /bDontSend=' +str(bDontSend)+ ' /AlreadySent=' + str(bAlreadySent) +']')
        continue
    
    if (strSingle == '') or (strVousTu == '') : 
        print ('line='+str(kk)+', ----> skip : ' + strInfos + ' / ' + strEmails + '[strSingle='+ str(strSingle)+ ' /strVousTu: ' +str(strVousTu)+ ']')
        continue

    strTimeNow = str(datetime.now().strftime('%d/%m/%Y %H:%M'))
    ## Add 4 min from the start time 
    # +' 06:00 PM'
    iMinutesToSend= iMinutesStart + (iToSend * iMinutesBetweenEmail)
    iSeconds = random.randrange(0,59,1)
    strMinutesSendAt =  Helpers.minutes2hour(iMinutesToSend,iSeconds)
    strToSendAt= strBaseSendTime + strMinutesSendAt
    #
    strEmailSubject = strBaseEmailSubject    
    if strType == 'TEST':
        strEmailSubject += ' (sent=['+ str(iToSend)+ ']/xl=['+ str(kk+1) +'], created=[' + strTimeNow + '],  to send=[' + strToSendAt +  '] )' 
    #
    strHTMLBody = Helpers.format_HTML_body(strSalut,isSingle,isVous,strAjout,strSignature,strSender,strType)
    if (strHTMLBody == '') or strHTMLBody.startswith('err'): 
        print ('line='+str(kk)+', ----> skip : ' + strInfos + ' / ' + strEmails + ', [strHTMLBody='+ strHTMLBody +']')
        continue
    if strEmails=='':
        print ('line='+str(kk)+', ----> skip : ' + strInfos + ' / ' + strSalut + ' / ' + strEmails + ', [no destinaitaire email]')
        continue
    #
    sent = sendmessage.send_mail_003(outlook,senderaccount,iShowDelayInSecs,strEmails,strEmailSubject,strHTMLBody,strType,strToSendAt,strTimeNow,kk,iMessagesSent)
    # Do not send tot many at once
    if sent :
        print ('line='+str(kk)+',>> SENT : ' + strInfos + ' / ' + strSalut + ' / ' +  strEmails + ', to send=[' + strToSendAt +  ']')
        iToSend += 1
        iMessagesSent += 1
        if iMessagesSent >= iMaxToSend :
            break
    else:
        iToSend += 1
        print ('line='+str(kk)+',==! SENT FAILED : ' + strInfos + ' / ' + strSalut + ' / ' +  strEmails + ' .')

    #

print (' ')
print ('Sending ended at ' + str(datetime.now()) + ', last kk is : ' + str(kk)+ ', mesges sent : ' + str(iMessagesSent))
print('''

::::::::::::::::::::::::::::::::::::::::::::::::::: ''')

