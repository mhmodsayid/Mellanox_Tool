import win32com.client as win32

outlook = win32.Dispatch('outlook.application')
#Data=outlook.mailItem
mail = outlook.CreateItem(0)
mail.To = 'mh_mouds@hotmail.com'
mail.Subject = 'Message subject'
mail.Body = 'Message body'
mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

print(mail.MessageClass)
print(mail.HTMLBody)

for account in outlook.Session.Accounts:
        if("@mellanox.com" in account.DisplayName):
            print(account.DisplayName)
           

# To attach a file to the email (optional):
#attachment  = "Path to the attachment"
#mail.Attachments.Add(attachment)

#mail.Send()