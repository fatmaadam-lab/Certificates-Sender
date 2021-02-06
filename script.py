'''
Certificates Sender
Note:
    before you gonna the program you have to :
    1- make sure to replace the certificate-template.docs with your certificate template 
    2- make sure to replace the trainer-data.xlsx with your data e-mails and names 
    3- in the send_certificate function that you replace the e-mail py your persnol email and your personal password
    4- maybe it required your permission to send the email via gmail so check your security permission in your gmail account

    thank you!
'''

import re
import xlrd
import email, smtplib, ssl
from datetime import datetime
from docx import Document
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



def docx_replace(regex, replace):
    '''
        this function got the selected words 
        and the replaced word then replaced 
        in word documents
    '''
    doc = Document('certificate-template.docx') #you can replace it by your own template
    for p in doc.paragraphs:
        if regex.search(p.text):
            inline = p.runs
            for i in range(len(inline)):
                if regex.search(inline[i].text):
                    text = regex.sub(replace, inline[i].text, count=1)
                    inline[i].text = text
                    doc.save('trainer-certificate.docx')
    return

def extract_xldr():
    '''
        this function extract all data 
        you need such as emails and names 
        and save it as array 
    '''
    wb = xlrd.open_workbook("trainer-data.xlsx") #you can replace it by your own data
    sheet = wb.sheet_by_index(0)
    new = [sheet.cell(0,cols).value for cols in range(sheet.ncols)]
    email_row = new.index('email')
    name_row  = new.index('Name  ')
    emails    = [sheet.cell(cols+1,email_row).value for cols in range(sheet.ncols+1)]
    names     = [sheet.cell(cols+1,name_row).value for cols in range(sheet.ncols+1)]
    return names,emails


def send_certificates():
    now = datetime.now()
    docx_replace(re.compile(r"Date"), now.strftime("%Y-%m-%d"))
    for a,b in zip(extract_xldr()[0],extract_xldr()[1]):
        docx_replace(re.compile(r"Student NAME"),a)

        # sending Email       
        subject = "pythonista, Congrats Your certificate is ready "
        body = "This is an email inform you completed succesfully our pythoncourse with your certificate has been sent as attachment"
        sender_email = 'email@gmail.com' #replace it with your email address 
        receiver_email = b
        password = 'password' #replace it with your password 

        # Create a multipart message and set headers
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject

        # Add body to email
        message.attach(MIMEText(body, "plain"))

        filename = "trainer-certificate.docx"  # In same directory as script

        # Open word docs file in binary mode
        with open(filename, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        # Encode file in ASCII characters to send by email    
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )

        # Add attachment to message and convert message to string
        message.attach(part)
        text = message.as_string()

        # Log in to server using secure context and send email
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, text)

        #send e-mail
    return 'has been sent successfully'

print(send_certificates())
