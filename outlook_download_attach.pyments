#libraries
import os,re,calendar,glob
from datetime import datetime, timedelta
import win32com.client

#Function 
def download_attachments(email_subject, output_folder):
    #Connecting to outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") 
    inbox = outlook.GetDefaultFolder(6) #This is for main inbox
    # If you need to access a specific subfolder, you'd need to modify the inbox selection: .Folders.Item("Folder Name")
    current_date = (datetime.now().date()) 
    #Looping through inbox items
    for item in inbox.Items:
        if item.Class == 43 and item.Subject == email_subject:  
           date_received = item.ReceivedTime
           date_str = date_received.strftime("%Y-%m-%d")
           updated_subject = f"{email_subject} {date_str}"
           folder_path = output_folder
           if date_str  == str(current_date):
                for attachment in item.Attachments:
                          attachment_filename, attachment_extension = os.path.splitext(attachment.FileName) 
                          attach_name = f"{attachment_filename} ({date_str}){attachment_extension}"
                          attachment.SaveAsFile(os.path.join(folder_path, attach_name))
                          print(f"Downloaded attachment: {attach_name}")
