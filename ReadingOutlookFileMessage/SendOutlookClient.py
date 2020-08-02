import win32com.client as win32
Outlook = win32.Dispatch('Outlook.application')
mail = Outlook.CreateItem(0)
mail.To = 'cesar.meneses@grupoprodeco.com.co'
mail.Subject = 'Message subject - Test'
mail.Body = 'Message body - Test'
mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

# To attach a file to the email (optional):
#attachment  = "Path to the attachment"
#mail.Attachments.Add(attachment)

mail.Send()