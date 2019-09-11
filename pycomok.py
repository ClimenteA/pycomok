
# coding: utf-8

# In[ ]:


import os, re
import datetime
from win32com.client import Dispatch


# In[ ]:


outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")


# In[ ]:


def get_email_address(mail_item):
    """Get full email address from the mail item"""
    
    if mail_item.Class==43:
        if mail_item.SenderEmailType=='EX':
            email_address = mail_item.Sender.GetExchangeUser().PrimarySmtpAddress
        else:
            email_address = mail_item.SenderEmailAddress
         
    elif mail_item.Class==4:
        try:
            email_address = mail_item.AddressEntry.GetExchangeUser().PrimarySmtpAddress
        except:
            return mail_item
    else:
        raise Exception("Can't extract email address from mail item!")
        
    return email_address


# In[ ]:


def process_filter_dates(starting_from_date, until_date): 
    """
        Process and return the dates given in outlook sql string format
    """
    #Date formats accepted with '-' and '/' as a separator
    d1 = "\d{2}-\d{2}-\d{4}"
    d2 = "\d{2}/\d{2}/\d{4}"

    if starting_from_date:
        
        starting_from_date_format_ok = False
        if re.search(d1, starting_from_date) or re.search(d2, starting_from_date):
            starting_from_date_format_ok = True
            
        if isinstance(starting_from_date, datetime.datetime):
            starting_from_date = starting_from_date.strftime('%d/%m/%Y')
        elif isinstance(starting_from_date, str) and starting_from_date_format_ok:
            starting_from_date = starting_from_date.replace("-", "/") 
        else:
            raise Exception("'starting_from_date' must be a datetime object or a string like '%d/%m/%Y'")

    if until_date:
        
        until_date_format_ok = False
        if re.search(d1, until_date) or re.search(d2, until_date):
            until_date_format_ok = True
                    
        if isinstance(until_date, datetime.datetime):
            until_date = until_date.strftime('%d/%m/%Y')
        elif isinstance(until_date, str) and until_date_format_ok:
            until_date = until_date.replace("-", "/") 
        else:
            raise Exception("'until_date' must be a datetime object or a string like '%d/%m/%Y'")
            
    return starting_from_date, until_date
    


# In[ ]:


def filter_items_by_date(mail_items, starting_from_date=None, until_date=None):
    """
        Filter folder items/messages by date
        'starting_from_date' - required will get all items from the starting date until now
        'until_date'         - optional will get all items until date specified
        NOT WORKING: 'received_time'      - the exact time when the mail was received like "29/08/2019 08:05"  
        If both arguments are present will get all items from 'starting_from_date' to 'until_date'
    """
    
    starting_from_date, until_date = process_filter_dates(starting_from_date, until_date)
    
    if starting_from_date == None and until_date == None:
        raise Exception("'starting_from_date' or 'until_date' needs to be completed!")

    if starting_from_date and until_date:
        print(f"Getting items starting from: {starting_from_date}, until date: {until_date}")
        mail_items = mail_items.restrict(f"[ReceivedTime] >= '{starting_from_date}' And [ReceivedTime] <= '{until_date}'")
   
    elif starting_from_date:
        print(f"Getting all items starting from: {starting_from_date}")
        mail_items = mail_items.restrict(f"[ReceivedTime] >= '{starting_from_date}'")
        
    elif until_date:
        print(f"Getting all items until date: {until_date}")
        mail_items = mail_items.restrict(f"[ReceivedTime] <= '{until_date}'")

    
    if mail_items.count == 0:
        raise Exception("No items match with the dates provided!")
    
    return mail_items


# In[ ]:


def get_item_recipients(mail_item):
    """
        Get from an email item the recipients/receivers in a list with dicts
    """
    receiversli = []
    for recipient in mail_item.Recipients:
        receiver = {
            "name": recipient.name,
            "email": get_email_address(recipient)
        } 

        receiversli.append(receiver)
            
    return receiversli


# In[ ]:


def get_items_data(mail_items, save_mail_item=False):
    """
        Get from Items object the data needed for later processing 
        Use 'filter_items_by_date' to get less data 
    """

    for item in mail_items:

        item_data = {
                 "received_date": item.ReceivedTime.strftime("%d/%m/%Y %H:%M"),
                 "subject": item.Subject,
                 "body": item.Body,
                 "HTMLBody": item.HTMLBody,
                 "sender": {"name": item.SenderName, "email": get_email_address(item)},
                 "receivers": get_item_recipients(item),
                 "MailItem": item if save_mail_item else "Outlook MailItem not saved set mail_item=True"
            }

        yield item_data


# In[ ]:


def get_mail_item(mail_items):
    """
        Get outlook MailItem object one by one (generator)
    """
    
    item = mail_items.GetFirst()
    while item:
        yield item
        item = mail_items.GetNext()


# In[ ]:


def get_accounts():
    """
        Get from outlook object all the accounts/emails set in outlook application
    """
    accounts_emails = [acc.SmtpAddress for acc in outlook.Accounts]
    accounts_names = [acc.CurrentUser.name for acc in outlook.Accounts]
    accounts = dict(zip(accounts_names, accounts_emails))
    
    return accounts


# In[ ]:


def get_outlook_mail_items(account_email=None, outlook_folder_path=None, display_accounts=False):
    """
        Get the outlook mail from a given folder (one_by_one generator func)
        account_email - in outlook you can set multiple mail addresses, 
        in account_email you must put the email which contains the folder needed
        outlook_folder_path - is the path to the folder you need to get items from
        outlook_folder_path must have this format "Inbox > foldername1 > foldername2" (folder names must be exact!)
        foldername2 - the last folder is where the items will be taken from
        By default it will take the first account found with the inbox folder
    """
    
    accounts = get_accounts()
    
    if display_accounts:
        print(accounts)
    
    if account_email == None and outlook_folder_path == None:
        account_email = accounts[list(accounts.keys())[0]]
        opfli = ["Inbox"]
        
    elif account_email != None and outlook_folder_path == None:
        opfli = ["Inbox"]
        
    elif account_email == None and outlook_folder_path != None:
        account_email = accounts[list(accounts.keys())[0]]
        opfli = outlook_folder_path.split(">")
        opfli = [f.strip() for f in opfli]
    
    else:
        opfli = outlook_folder_path.split(">")
        opfli = [f.strip() for f in opfli]
        

    account_folders = outlook.Folders.Item(account_email)
    
    items_expli = ["account_folders"]
    for outlook_folder in opfli:
        items_expli.append(f"Folders('{outlook_folder}')")
    
    
    exp = ".".join(items_expli) + ".Items"
    
    if display_accounts:
        print(exp)
    
    try:
        
        items_com_object = eval(exp)
        return items_com_object
    
    except Exception as e:
        print(str(e))
        errmsg = f"Can't get 'Items' from account_email: {account_email} with this outlook_folder_path: {outlook_folder_path}"
        raise Exception(errmsg)


# In[ ]:


def get_to_cc(mails):
    """
        Check if the format of the mails and return the mail in To, CC string format
    """
    if isinstance(mails, list):
        mails = "; ".join(mails)
    elif not isinstance(mails, str):
        raise Exception("Must be str or list!")
        
    return mails


# In[ ]:


def send_email(subject, message, to, cc=None, attachments=None, display=True, send=True):
    """
        Send an email using the default email profile
        subject     - string representing the subjet of the mail, 
        message     - string representing the body of the mail, 
        to          - string with an email or a list of emails, 
        cc          - string with an email or a list of emails, 
        attachments - string with a file path or a list of file paths, 
        display     - True show email before sending False otherwise
    """
    #TODO - the email is send from the default email, with logon should be sent from the mail selected
    # Dispatch("Mapi.Session").Logon('ACCRT_LECT_BFBF-av')

    new_mail = Dispatch("Outlook.Application").CreateItem(0)

    new_mail.Subject = str(subject)
    new_mail.Body = str(message)

    new_mail.To = get_to_cc(to)
    
    if cc:
        new_mail.CC = get_to_cc(cc)

    
    if not isinstance(attachments, list):
        raise Exception("'attachments' must be a list!")
        
    for filepath in attachments:
        
        if not os.path.isfile(filepath):
            raise Exception(f"{filepath} not found!")
        
        new_mail.Attachments.Add(filepath)
    
    if display:
        new_mail.Display(True)
    
    if send:
        new_mail.Send()


# In[ ]:


# from pycomok import get_outlook_mail_items, filter_items_by_date, get_items_data, send_email, get_mail_item

# items = get_outlook_mail_items('alincmt@gmail.com', "Inbox > planning", True)

# mail_items = filter_items_by_date(items, starting_from_date="04-09-2019", until_date=None)

# gen_items = get_items_data(mail_items)

# for item_data in gen_items:
#     break

# item_data['subject']

# item_data['body']

#If get_items_data(mail_items, save_mail_item=True)
# item_data['MailItem'].SaveAs("absolute/path/name_of_file.msg")



# item_generator = get_mail_item(bf_items)

# i=0
# for item in item_generator:
#     i += 1

