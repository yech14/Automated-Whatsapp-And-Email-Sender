

# Project Automated Mail Sender
# Sendings emails to a list of emails automatically

## To Do List:
# -


import smtplib, ssl
import time
from email.message import EmailMessage
import pandas as pd
import random

# Load the data frame
df = pd.read_excel('C:/Users/.../mails_of.xlsx')
mails = list(df['mail:'])
names = list(df['name:'])

# The setting for the server
port = 465  # For SSL
smtp_server = "smtp.gmail.com"
sender_email = 'your email address'  # הכנס את כתובת המייל שלך כאן
password = "your password or token"  # הכנס את הסיסמא או הטוקן שלך כאן

# Make subgroup for test
names1 = names[:]
receiver_emails = mails[:]


for i, name in enumerate(names1):
    # Set up message body and signature
    message_body = f"""\
<html>
<body>
<pre style="font-family: Arial, sans-serif;">
<div dir="rtl">
Hello {name}, 
 
write your message here..

</div>
</pre>

</body>
</html>

"""
    signature = 'your signature here' # הכנס את החתימה שלך כאן

    # Create message object
    msg = EmailMessage()
    msg.set_content(message_body + signature, subtype='html')
    msg['From'] = "your name" # הכנס את השם שלך כאן
    msg['Subject'] = "your subject here" # הכנס את הנושא כאן

    # Iterate through the list of email addresses and send the message to each one

    msg['To'] = receiver_emails[i]
    # Connect to SMTP server and send email
    context = ssl.create_default_context()

    time.sleep(random.uniform(1, 7))

    with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
        server.login(sender_email, password)
        server.send_message(msg, from_addr=sender_email, to_addrs=receiver_emails[i])
