import win32com.client as win32
Outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
#accounts = win32.Dispatch("Outlook.Application").Session.Accounts
inbox = Outlook.Folders('helpdesk@grupoprodeco.com.co').Folders('Bandeja de entrada')
# List folder in inbox.
#inbox1 = Outlook.GetDefaultFolder(6).Folders.Item("Your_Folder_Name")
message = inbox.Items.GetLast()

print(message.Sender)
print(message.SenderName)
print(message.To)
print(message.Sender.Address)

#print(message.Subject)
#print(message.Categories)
#print(message.SenderEmailAddress)
#print(message.Body)

#print(messages.Subject)
#print(messages.Categories)
#print(messages.SenderEmailAddress)
#print(messages.Body)