import datetime
import os
import win32com.client


path = os.path.expanduser('C:\Workspace\Outlook_Attachments\Attachments')

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace('MAPI')
# finds messages in the 'inbox' folder
inbox = outlook.GetDefaultFolder(6)
#finds messages in the 'sent' folder
sent = outlook.GetDefaultFolder(5)

#takes files from specified folder and saves the attachments to path
def get_files():
    inbox_messages = inbox.items
    sent_messages = sent.items
    saveattachemnts(inbox_messages)
    saveattachemnts(sent_messages)

def saveattachemnts(messages):
    for message in messages:
        attachments = message.Attachments
        if message.Class == 43 and attachments.Count != 0:
            try:
                attachment = attachments.Item(1)
                for attachment in message.Attachments:
                    if attachment.FileName[-4:] != ".png" and attachment.FileName[-4:] != ".gif" :
                        datecode = attachments.Parent.SentOn
                        attachment.SaveAsFile(os.path.join(path, f"{attachment}"))
                        print(f"Processed {message}")
                        break
                    else:
                        continue
            except Exception as e:
                print(f'Error with {message}.')
        else:
            print("No Attachments")
            continue
