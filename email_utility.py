"""
Email utility for CCC Tools Website
Author: Franz Weishaupl
"""

import smtplib
#from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.base import MIMEBase


# make into a proper class
class Email_Utility:
    """
Email utility for CCC Tools Website
Author: Franz Weishaupl

(ARGS 1, 2, 3)  Supply a LIST of "to", "cc" & "bcc" recipients
(ARGS 4, 5)     Supply subject & body text
(ARGS 6, 7)     Supply optional file & name of file with filetype
    
    """
    
    def __init__(self, to, cc, bcc, subject, body, file, filename):
        self.to = to
        self.cc = cc
        self.bcc = bcc
        self.subject = subject
        self.body = body
        self.file = file
        self.filename = filename
        
    def Send(self):
        '''
        Send email
        '''
        msg = MIMEMultipart()
        msg['To'] = ", ".join(self.to)
        msg['Cc'] = ", ".join(self.cc)
        msg['Bcc'] = ", ".join(self.bcc)
        msg['From'] = 'F1800164@team.telstra.com'
        msg['Subject'] = self.subject
        text_body = MIMEText(f'{self.body}')
        attachment = MIMEMultipart()
        attachment = MIMEBase('application', "octet-stream")
        
        # Attempt to attach file
        try:
            attachment.set_payload(open(self.file, "rb").read())        
            encoders.encode_base64(attachment)
            attachment.add_header('Content-Disposition', f'attachment; filename={self.filename}')
            msg.attach(attachment)
        except FileNotFoundError as e:
            print(e)
        
        msg.attach(text_body)
        
        try:
            server = smtplib.SMTP('smtp.smart-rr.in.telstra.com.au')
            server.sendmail(msg['From'], [msg['To'], msg['Cc'], msg['Bcc']], msg.as_string())
            server.close()
            print('Email sent!')
        except Exception as e:
            print('Something went wrong...')
            print(e)
    
