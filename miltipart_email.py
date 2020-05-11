import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import csv


def send_mail(SUBJECT, BODY, TO, FROM):
    """With this function we send out our html email"""
     # Create message container - the correct MIME type is multipart/alternative here!
    MESSAGE = MIMEMultipart('alternative')
    MESSAGE['subject'] = SUBJECT
    MESSAGE['To'] = TO
    MESSAGE['From'] = FROM
    MESSAGE.preamble = ""

    # the MIME type text/html.
    HTML_BODY = MIMEText(BODY, 'html')  # Attach parts into message container.
    # According to RFC 2046, the last part of a multipart message, in this case
    # the HTML message, is best and preferred.
    MESSAGE.attach(HTML_BODY)  # The actual sending of the e-mail
    # Print debugging output when testing
    server = smtplib.SMTP('smtp.gmail.com:587')

    # server.set_debuglevel(1)  # Credentials (if needed) for sending the mail
    password = ""
    server.starttls()
    server.login(FROM, password)
    try:
        # server.sendmail(FROM, [TO], MESSAGE.as_string())
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

if __name__ == "__main__":
    """Executes if the script is run as main script (for testing purposes)"""
    email_content = """<meta http-equiv="Content-Type" content="text/html; charset=utf-8"> <title>Notification</title>
         <style type="text/css" media="screen">
         table{
         background-color: #AAD373;
         empty-cells:hide;
         }
         td.cell{
         background-color: white;
         }
         </style>  <table style="border: blue 1px solid;">  <tbody>
        <tr>
            <td class="cell">Cell 1.1</td>
            <td class="cell">Cell 1.2</td>
        </tr>  <tr>
            <td class="cell">Cell 2.1</td>
            <td class="cell"></td>
        </tr>  </tbody>
        </table>  """

    data = read_csv()
    FROM ='xyz@gmail.com'
    for email_data in data:
        send_mail("Test email subject", email_content, email_data['email'], FROM)
