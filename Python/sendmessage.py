from win32com.client import Dispatch
import win32com
import time


def init_ol(strSender):
    outlook = win32com.client.Dispatch("Outlook.Application")
    for accoun in outlook.Session.Accounts:
        if accoun.SmtpAddress == strSender : ## 'your@mail.com':
            senderaccount = accoun
            break
    return outlook,senderaccount


def send_mail_003(outlook,senderaccount,iShowDelayInSecs, strEmails,strEmailSubject,strHTMLBody,strType,strSendTime,strNow,iListId,iOrdInThisRun):

    mail = outlook.CreateItem(0)
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, senderaccount))
    ### parse emails destinataion(s)

    if strEmails=='':
            return False
    strEmails = strEmails.replace(' ',';',)
    strEmails = strEmails.replace(',',';')
    while strEmails.find(';;') >= 0:
          strEmails = strEmails.replace(';;',';')
    strEmailTo = strEmails.split(';')
    if len(strEmailTo) == 1:
        mail.To = strEmailTo[0]
    else:
        for strEmail in strEmailTo:
             mail.Recipients.Add(strEmail)
    #........

    mail.Subject = strEmailSubject

    mail.BodyFormat = 2   # 2: Html format // olFormatHTML
    mail.HTMLBody = strHTMLBody

    ### Prob redundant
    # mail_item.InternetCodepage=28591 ## iso-8859-1 	28591
    ## properties = mail.ItemProperties() 
    ## thisProperty =properties(0)
                ## MailIsDelayed = False               ' We assume it's being delivered now
                ## NoDeferredDelivery = "1/1/4501"     ' Magic number Outlook uses for "delay mail box isn't checked"
    ## print(thisProperty.value) 
    #

    # mail.Attachments.Add(att)
    mail.DeferredDeliveryTime = strSendTime
    if iShowDelayInSecs > 0:
        mail.Display(False) ## Not modal
        time.sleep(iShowDelayInSecs)
    mail.Save
    mail.Close(0)   # olSave = 0
    ## 
    # mail.SaveAs(9) # Additinal save, not necessary. OlSaveAsType=9 is Format de message unicode Outlook (.msg)
    ### make sure it is saved :
    if not (mail.Saved): 
       mail.Save
    ###
    mail.send
    return True


