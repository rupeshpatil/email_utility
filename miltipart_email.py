import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.generator import Generator
import csv
import  pandas as pd

import jinja2


def send_mail(SUBJECT, TO, CC,FROM):
    """With this function we send out our html email"""
     # Create message container - the correct MIME type is multipart/alternative here!
    print("to",TO)
    print("CC",CC)
    MESSAGE = MIMEMultipart('alternative')
    MESSAGE['subject'] = SUBJECT
    MESSAGE['To'] = TO
    MESSAGE['From'] = FROM
    MESSAGE['cc'] = CC
    MESSAGE.preamble = ""

    # the MIME type text/html.
    # HTML_BODY = MIMEText(BODY, 'html') 
    # HTML_BODY = MIMEText(
    #        Environment().from_string(TEMPLATE).render(
    #            participant=TO
    #        ), "html"
    #    )
   
    # Load template from html file
    templateLoader = jinja2.FileSystemLoader(searchpath="./")
    templateEnv = jinja2.Environment(loader=templateLoader)
    TEMPLATE_FILE = "email_template.html"
    template = templateEnv.get_template(TEMPLATE_FILE)
    outputText = template.render(participant= TO) 
    
    HTML_BODY = MIMEText(outputText, "html")
    # Attach parts into message container.
    # According to RFC 2046, the last part of a multipart message, in this case
    # the HTML message, is best and preferred.
    MESSAGE.attach(HTML_BODY)  # The actual sending of the e-mail
    # Print debugging output when testing
    server = smtplib.SMTP('smtp.gmail.com:587')
    server = smtplib.SMTP('smtp-mail.outlook.com:587')

    # server.set_debuglevel(1)  # Credentials (if needed) for sending the mail
    password = "xyz"
    server.starttls()
    server.login(FROM, password)
    try:
       
        server.sendmail(FROM, [TO], MESSAGE.as_string())
        print('Email to {} successfully sent!\n\n'.format(TO))
    except Exception as e:
        print('Email to {} could not be sent :( because {}\n\n'.format(TO, str(e)))

    server.quit()

# Read comma separated CSV
def read_csv():
    with open('employee_notification.csv', mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        for data in csv_reader:
            yield data
            
def read_excel():
   
    excel_data_df = pd.read_excel('data.xlsx', sheet_name='Foundation pipeline (2)')
    email_lst = excel_data_df['User Email'].tolist()
    return email_lst

def read_xlsb():
    df = pd.read_excel('GAD12.xlsx')
    # dt = df[['Email ID','Supervisor Email ID']].to_dict('records')
    dt = df[['Email ID','Supervisor Email ID']]
    data_dict = dt.set_index('Email ID')['Supervisor Email ID'].to_dict()
    return data_dict

    
if __name__ == "__main__":
    """Executes if the script is run as main script (for testing purposes)"""
    
    data = read_excel()
    supervisor_data = read_xlsb()
    FROM ='xyz@gmail.com'
    for email in data:
        send_mail("Automation Test Email", email, supervisor_data.get(email,''), FROM)
