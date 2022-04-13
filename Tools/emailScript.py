import shutil
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time


port = 465
smtp_server = "smtp.gmail.com"
sender_email = "tdmUsageAlert@gmail.com"
recipients = ["DDiaz@tradedatamonitor.com"]
password = "tdm12345$"     #tdm12345$

message = MIMEMultipart("alternative")
message["Subject"] = 'C Drive Capacity'
message["from"] = "TDM Data Team"
message["To"] = ", ".join(recipients)

text = """\
The C drive has exceeded 90% capacity
Path:\t \\\\192.168.100.10\\e$
"""
html = """\
<html>
    <body>
      <p>The C drive has exceeded 90% capacity<br>
         Path:\t \\\\192.168.100.10\\c$
      <p>
    </body>
  </html>
  """
part1 = MIMEText(text, "plain")
part2 = MIMEText(html, "html")

message.attach(part1)
message.attach(part2)

context = ssl.create_default_context()

with smtplib.SMTP_SSL("smtp.gmail.com",port, context = context) as server:
    server.login("tdmUsageAlert@gmail.com", password)
    server.sendmail(sender_email, recipients, message.as_string())




"""
To set up a Gmail address for testing your code, do the following:

Create a new Google account.
Turn Allow less secure apps to ON. Be aware that this makes it easier for others to gain access to your account.
"""
