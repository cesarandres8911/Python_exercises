# https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem
# https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem.body
# https://www.thetopsites.net/article/51621535.shtml
# https://codereview.stackexchange.com/questions/213660/outlook-email-rules-and-attachment-downloader-in-python
# https://www.codetd.com/en/article/8789029
# https://stackoverflow.com/questions/22813814/clearly-documented-reading-of-emails-functionality-with-python-win32com-outlook
# https://readthedocs.org/projects/pyoutlook/downloads/pdf/stable/

import win32com.client as win32
Outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = Outlook.Folders('helpdesk@grupoprodeco.com.co').Folders('Bandeja de entrada')
message = inbox.Items.GetLast() # The last item
folder = inbox.Items
for message in folder:
    if message.Unread == True:
        print ("Esta como no leído")
        print(message.SenderName)
        print(message.To)
        print(message.Sender.Address)
        print(message.Subject)
        print(message.Categories)
        print(message.SenderEmailAddress)
        print(message.Body)