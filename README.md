# pycomok
Access outlook emails using win32com!
https://docs.microsoft.com/en-us/office/vba/api/overview/outlook

module is still in testing

How to:

```
from pycomok import get_outlook_mail_items, filter_items_by_date, get_items_data, send_email, get_mail_item

items = get_outlook_mail_items('alincmt@gmail.com', "Inbox > planning", True)

mail_items = filter_items_by_date(items, starting_from_date="04-09-2019", until_date=None)

gen_items = get_items_data(mail_items)

for item_data in gen_items:
     break

item_data['subject']

item_data['body']

#If get_items_data(mail_items, save_mail_item=True)
item_data['MailItem'].SaveAs("absolute/path/name_of_file.msg")



item_generator = get_mail_item(bf_items)

i=0
for item in item_generator:
     i += 1
```
