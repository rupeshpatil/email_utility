from smtplib import SMTP              # sending email
from email.mime.text import MIMEText  # constructing messages

from jinja2 import Environment        # Jinja2 templating

TEMPLATE = """
<html>
<head>
</head>
<body>
<p class="MsoNormal">
    <a name="_Hlk496010796"> Dear {{participant}},,
        <o:p></o:p>
    </a>
</p>
<p class="MsoNormal">
This is for Automation Academy Email utility testing.
</p>

Please disregard


</body>
</html>
"""  # Our HTML Template

subject = "Automation Test Email"
sender= "xyz@gmail.com"
recipient = "xyz@gmail.com"

# Create a text/html message from a rendered template
msg = MIMEText(
    Environment().from_string(TEMPLATE).render(
        participant= recipient
    ), "html"
)



msg['Subject'] = subject
msg['From'] = sender
msg['To'] = recipient

# server = smtplib.SMTP('smtp.gmail.com:587')
server = SMTP('smtp-mail.outlook.com:587')

# server.set_debuglevel(1)  # Credentials (if needed) for sending the mail
password = ""
server.starttls()
server.login("xyz@gmail", password)

try:
    server.sendmail(sender, [recipient], msg.as_string())
    print('Email to {} successfully sent!\n\n'.format(recipient))
except Exception as e:
    print('Email to {} could not be sent :( because {}\n\n'.format(recipient, str(e)))

server.quit()# -*- coding: utf-8 -*-

