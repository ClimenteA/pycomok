# pycomok
Access outlook emails using win32com!
https://docs.microsoft.com/en-us/office/vba/api/overview/outlook


Short how to:

Import the needed funcs (not many)

```
from pycomok import get_outlook_mail_items, filter_items_by_date, get_items_data, send_email, get_mail_item
```

#Get the mail items from 'planning' folder
```
email = 'alincmt@gmail.com' # put the email address from which you want to get the mails
mail_items_folder_path = "Inbox > planning" # here is the path to 'planning' outlook folder, use '>' for path 
items = get_outlook_mail_items(email, mail_items_folder_path, True)
```
Filter mail items by date (recomended if to many)
```
mail_items = filter_items_by_date(items, starting_from_date="04-09-2019", until_date=None)
```
Make a generator from the mail_items filtered and iterete over them to get data
```
gen_items = get_items_data(mail_items)

for item_data in gen_items:
     break
```
Investigate the data retreived
```
item_data['subject']
>>> 'meeting at 4pm'
item_data['body']
>>> 'hello there, we have a meeting at 4pm blabla'
```

The function 'get_items_data' has 'save_mail_item=False' by default if you set it to True you can access the outlook 'MailItem' object
```
item_data['MailItem'].SaveAs("absolute/path/name_of_file.msg")
```
Here is the item generator
```
item_generator = get_mail_item(bf_items)

i=0
for item in item_generator:
     i += 1
```

It's just one file ('pycomok.py' aka python + win32com + outlook)
