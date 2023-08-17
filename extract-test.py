import imaplib
import email

import yaml  

with open("credentials.yml") as f:
    content = f.read()

my_credentials = yaml.load(content, Loader=yaml.FullLoader)

user, password = my_credentials["user"], my_credentials["password"]

imap_url = 'imap.gmail.com'


my_mail = imaplib.IMAP4_SSL(imap_url)

my_mail.login(user, password)

my_mail.select('Inbox')


key = 'FROM'
value = 'priscilla.pavan@videosoft.com.br'
_, data = my_mail.search(None, key, value)  

mail_id_list = data[0].split() 

msgs = [] 

for num in mail_id_list:
    typ, data = my_mail.fetch(num, '(RFC822)') 
    msgs.append(data)


email_results = []

for msg in msgs[::-1]:
    for response_part in msg:
        if type(response_part) is tuple:
            my_msg = email.message_from_bytes((response_part[1]))
            email_info = {
                "subject": my_msg['subject'],
                "from": my_msg['from'],
                "body": ""
            }
            for part in my_msg.walk():  
                if part.get_content_type() == 'text/plain':
                    email_info["body"] = part.get_payload()
            email_results.append(email_info)
