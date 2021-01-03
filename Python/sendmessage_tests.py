import win32com.client as win32



def send_mail_001():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'shalev@vayness.com'
    mail.Subject = 'Message subject'
    mail.Body = 'Message body'
    mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

    # To attach a file to the email (optional):
    attachment  = "Path to the attachment"
    mail.Attachments.Add(attachment)

    mail.Send()
    return True


def send_mail_002():
    outlook_app = win32.Dispatch('Outlook.Application')

    # choose sender account
    send_account = None
    for account in outlook_app.Session.Accounts:
        if account.DisplayName == 'shalev@isako.com':
            send_account = account
            break

    mail_item = outlook_app.CreateItem(0)   # 0: olMailItem

    # mail_item.SendUsingAccount = send_account not working
    # the following statement performs the function instead
    mail_item._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

    mail_item.Recipients.Add('receipient@outlook.com')
    mail_item.Subject = 'Test sending using particular account'
    mail_item.BodyFormat = 2   # 2: Html format
    mail_item.HTMLBody = '''
        <H2>Hello, This is a test mail.</H2>
        Hello Guys. 
        '''
    mail_item.Send()        
    return True
