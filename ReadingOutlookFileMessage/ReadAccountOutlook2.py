import win32com.client as win32
Outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = Outlook.Folders('helpdesk@grupoprodeco.com.co').Folders('Bandeja de entrada')

messages = inbox.Items
message = inbox.Items.GetFirst()
i = 0
while message:
    if message.Unread == True and message.Categories == "Pendiente":
        print (message.Subject)
        print(message.SenderName)
        print(message.To)
        print(message.Sender.Address)
        print(message.Subject)
        print(message.Categories)
        print(message.SenderEmailAddress)
        print(message.Body)
    message = messages.GetNext()