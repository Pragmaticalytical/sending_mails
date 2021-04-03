"""
********************************************************************
NAME: pragma_mail/send_mail.py
DESCRIPTION: Sends mail to recipients with attachments
DATE LAST MODIFIED: 2021/03

AUTHOR: Marcel Coetzee - https://pragmaticalytical.com/
********************************************************************
"""

import atexit
import configparser
import os
import smtplib
import ssl
from email.encoders import encode_base64
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

if os.name == "nt":
    try:
        import pywintypes
        import win32com.client as local_outlook_mail_client
    except ImportError:
        raise Exception("Please install pywintypes by running 'pip install pypiwin32'")

config_parse = configparser.ConfigParser()


class mailer(object):
    global config_parse

    def __init__(
        self, client_type="local_outlook", smtp_config_file="smtp.ini"
    ) -> None:
        super().__init__()

        if client_type in ["local_outlook", "remote_server"]:

            self.client_type = client_type

            if self.client_type == "local_outlook":
                try:
                    self.outlook_object = local_outlook_mail_client.Dispatch(
                        "outlook.application"
                    )
                    atexit.register(self.cleanup)
                except pywintypes.com_error:

                    ## From https://stackoverflow.com/questions/41611383/how-to-connect-to-a-running-instance-of-outlook-from-python:
                    # Outlook is a singleton, so no matter what you do, it will always connect to the running instance.
                    # The only problem (as you have discovered) is that if is already running under a different security context,
                    # COM system will refuse to marshal COM objects between the two processes.

                    raise RuntimeError(
                        "Please close Outlook before running this script!\nRefer to https://stackoverflow.com/questions/41611383/how-to-connect-to-a-running-instance-of-outlook-from-python"
                    )

            elif self.client_type == "remote_server":

                try:
                    config_parse.read(smtp_config_file)
                except Exception as e:
                    raise Exception("No 'smtp.ini' config file found!")

                self.SERVER = config_parse["SMTP_SERVER"]["SERVER"]
                self.port = config_parse["SMTP_SERVER"]["port"]
                self.user = config_parse["SMTP_SERVER"]["user"]
                self.password = config_parse["SMTP_SERVER"]["password"]

                self.context = ssl.create_default_context()

                self.server_instance = smtplib.SMTP_SSL(
                    self.SERVER, self.port, context=self.context
                )

                self.server_instance.login(self.user, self.password)

                # to close the server connection at exit:
                atexit.register(self.cleanup)

        else:
            raise NotImplementedError(
                f"{client_type} isn't implemented yet.\nPlease choose one of either 'local_outlook' or 'remote_server'"
            )

    def cleanup(self):
        try:
            self.SERVER.quit()  # https://stackoverflow.com/questions/3850261/doing-something-before-program-exit
            self.outlook_object.quit()
        except:
            pass

    def send_mail(
        self,
        sender_mail: str = "",
        recipients: list[str] = [""],
        cc_recipients: list[str] = [""],
        mail_subject: str = "",
        mail_body: str = "",
        attachment_path: list[str] = None,
        rat_pic: bool= False,
    ):

        if self.client_type == "local_outlook":

            mail_instance = self.outlook_object.CreateItem(0)
            mail_instance.To = ("; ").join(recipients)

            mail_instance.Subject = mail_subject
            mail_instance.HTMLBody = mail_body

            # if attachment_path is not None:
            #     for file in attachment_path:
            #         mail_instance.Attachments.Add(file)

            if attachment_path is not None:
                for file in attachment_path:
                    mail_instance.Attachments.Add(file)

            mail_instance.Send()
            print("Mail successfully sent.")

        elif self.client_type == "remote_server":

            msg = MIMEMultipart("alternative")

            ##-- Mail body --##
            body = MIMEText(mail_body, "html")
            # msg.add_header("Content-Type", "text/html")
            msg.attach(body)

            ##-- Mail body --##

            ##-- File attachments --##
            # https://stackoverflow.com/questions/2798470/binary-file-email-attachment-problem

            if attachment_path is not None:
                for file in attachment_path:

                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(open(file, "rb").read())
                    encode_base64(part)
                    part.add_header(
                        "Content-Disposition", f"attachment; filename={file}"
                    )
                    msg.attach(part)
            ##-- File attachments --##

            ## -- rat picture -- ## 
            # https://stackoverflow.com/questions/60430283/how-to-add-image-to-email-body-python

            if rat_pic:

                rat_image_location = "./rat.jpg"

                with open(rat_image_location, "rb") as fp:
                    img = MIMEImage(fp.read())
                    img.add_header("Content-ID", "<{}>".format("rat"))
                    msg.attach(img)

            ## -- rat picture -- ##


            sender = self.user
            receivers = recipients

            msg["Subject"] = mail_subject
            msg["From"] = sender_mail
            msg["To"] = (", ").join(recipients)
            msg["Cc"] = (", ").join(cc_recipients)

            self.server_instance.sendmail(sender, receivers, msg.as_string())

            print("mail successfully sent")
