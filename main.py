"""
********************************************************************
NAME: main.py
DESCRIPTION: Sends mail to recipients with attachments, utilizing pragma_mail
DATE LAST MODIFIED: 2021/03

AUTHOR: Marcel Coetzee - https://pragmaticalytical.com/
********************************************************************
"""
import os
import re
from pathlib import Path
from jinja2 import Template

from pragma_mail import mail

## Configuration ##
client_type = "local_outlook"  # Either Remote SMTP server ("remote_server"), or Outlook application on your machine ("local_outlook")
smtp_config_file = "./smtp.ini"

## Senders, Recipients, Subject line ##
sender_mail = "yourmail@yourdomain.com"
recipients = ["recipient1@recipientdomain.com", "recipient3@recipientdomain.com"]
cc_recipients = ["recipient3@recipientdomain.com", "recipient4@recipientdomain.com"]
mail_subject = "Mail from rat"  # <- Subject line text

## Attachments ##
files = ["rat.png", "super_legit_worksheet.xlsx"]

files_path_base = Path(os.getcwd())
attachments_list = [files_path_base / file for file in files]
attachments_list = [str(file_path) for file_path in attachments_list]

## Mail body template ##
recipients_names_search = [
    re.search("^.*\@", recip, re.IGNORECASE) for recip in recipients
]
to_names = [name.group().upper()[:-1] for name in recipients_names_search if name]

# templating, you can just replace the below with plain strings instead of this fancy regex
# to_names = (", ").join(to_names)
to_names = (", ").join(to_names)
num_rep = str(len(files))
my_name = re.search(r"^.*\@", sender_mail, re.IGNORECASE).group().upper()[:-1]

with open("./assets/mail_template.html", "r") as mail_template_file:
    mail_body = Template(mail_template_file.read())
    mail_body = mail_body.render(to_name=to_names, num_rep=num_rep, my_name=my_name)


if __name__ == "__main__":

    mailer = mail.mailer(client_type=client_type, smtp_config_file=smtp_config_file)
    mailer.send_mail(
        sender_mail=sender_mail,
        recipients=recipients,
        cc_recipients=cc_recipients,
        mail_subject=mail_subject,
        mail_body=mail_body,
        attachment_path=attachments_list,
        rat_pic=True
    )
