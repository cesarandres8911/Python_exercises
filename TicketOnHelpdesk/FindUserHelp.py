import win32com.client


outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
msg = outlook.OpenSharedItem(r'C:\temp\test\codigo acceso TEAMS.msg')

username = msg.SenderEmailAddress[msg.SenderEmailAddress.rfind("CN=")+3:len(msg.SenderEmailAddress)]
#r'C:\Recursos\Programacion\Python\Python_exercises\TicketOnHelpdesk\ActiveDirectory.xls'


