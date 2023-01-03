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


iMaxToSend = 6 # 40

# Filtrage avec Prio (from 01/2022)
iPrioMinForThisRun=10  #1 and up, 0=ignore
iPrioMaxForThisRun=10  #1 and up, 0=ignore, 'Max' >= 'Min'
if (iPrioMaxForThisRun < iPrioMinForThisRun):
    # Stop 
    print (' ***  STOP STOP :: iPrioMaxForThisRun ('+ str(iPrioMaxForThisRun) +') < iPrioMinForThisRun ('+ str(iPrioMinForThisRun) +')')
    quit()

dtNowAtStart=datetime.now()
#. . . . . . . . . 
# # make sure the right line is the last of the following 2
bSendToday = True
if bSendToday:
    strBaseSendDate = dtNowAtStart.strftime('%d/%m/%Y')
else:
    strBaseSendDate='03/01/2023' ## If a predetermined date
# make sure explicit date is later than today :
if strBaseSendDate < dtNowAtStart.strftime('%d/%m/%Y'):
    # Stop 
    print (" ***  STOP STOP :: attempting to send at a past date ...")
    quit()
     

#. . . . . . . . . 
# make sure the right line is the last of the following 2
bSendNow = True
if bSendNow:
    iMinutesStart =int(dtNowAtStart.strftime('%H'))*60 + int(dtNowAtStart.strftime('%M')) ## "%H:%M:%S.%fZ"
else:
    iMinsPerHour = 60
    iMinutesStart = (11 * iMinsPerHour) + 28 # If a predetermined time, here rounded to an hour, so '(8 * 60) + 25 ' is start sending at 8h25
    # If the Base date is today than make sure we're not earlier than now :
    if bSendToday:
        iMinutesNow =int(dtNowAtStart.strftime('%H'))*60 + int(dtNowAtStart.strftime('%M')) ## "%H:%M:%S.%fZ"
        iMinutesStart = max(iMinutesNow,iMinutesStart)

# So first mail is not earler than 5 mins from now 
iMinutesStart += 5 
#


# print (proximate value, car seconds are set to '00')
# strBaseSendTime =  Helpers.minutes2hour(iMinutesStart,0)
### NOK :: strBaseSendTime =str(dtNowAtStart.strftime('%H:%M'))

# New sheet every year
strFile='..\..\HNY_Emails_MultiYears.xlsx'
strSheet='HNY_2023'
strBaseEmailSubject='Belle et heureuse année 2023 !' 
iMinutesBetweenEmail = 3 ## 
iShowDelayInSecs = 3
strTypesToSend = '__,_TEST_' # The '__' means it's OK to send where the TYPE Column is empty, 
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

print('base start date-time is ' + strBaseSendDate + ' ' + strTime + ', time now at start of round : ' + str(dtNowAtStart.strftime('%d/%m/%Y %H:%M')) + ', (' + strBaseSendDate +')')
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
    iPrio = 0 if pd.isnull(sheet['PRIO'][kk]) else int(sheet['PRIO'][kk]) # 

    bTypeNotOK = True if strTypes_NOT_ToSend.find('_'+strType+'_')>=0 else False
    bTypeOK = True if strTypesToSend.find('_'+strType+'_')>=0 else False

    # Added 01/2022 - premier filter using Prio
    if (iPrioMinForThisRun >= 1) and (iPrioMinForThisRun < iPrio): 
        print ('line='+str(kk)+', ----> skip : ' + strInfos + ' / ' + strEmails + '[iPrioMinForThisRun='+ str(iPrioMinForThisRun)+ ' == 0 or < ' +str(iPrio) +']')
        continue
    if (iPrioMaxForThisRun >= 1) and (iPrioMaxForThisRun > iPrio): 
        print ('line='+str(kk)+', ----> skip : ' + strInfos + ' / ' + strEmails + '[iPrioMaxForThisRun='+ str(iPrioMaxForThisRun)+ ' == 0 or > ' +str(iPrio) +']')
        continue

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
    ### NOK :: strToSendAt= strBaseSendTime + strMinutesSendAt
    strToSendAt= strBaseSendDate + strMinutesSendAt
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
print ('Sending ended at ' + str(datetime.now()) + ', last kk is : ' + str(kk)+ ', msges sent : ' + str(iMessagesSent))
print('''

::::::::::::::::::::::::::::::::::::::::::::::::::: ''')

