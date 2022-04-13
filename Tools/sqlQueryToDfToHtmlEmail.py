import pandas as pd
import pyodbc
import shutil
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
from IPython.display import display, HTML

def emailTable(df):
    port = 465
    smtp_server = "smtp.gmail.com"
    sender_email = "tdmUsageAlert@gmail.com"
    recipients = ["DDiaz@tradedatamonitor.com"]
    password = "tdm12345$"     #tdm12345$

    message = MIMEMultipart("alternative")
    message["Subject"] = '[EN] HS Fields Starting with Dashes'
    message["from"] = "TDM Data Team"
    message["To"] = ", ".join(recipients)

    html = df.to_html(index=False, classes='table table-striped')
    # print(html)

    html = f"""\
    <html>
        <body>
            <p>[HSDESC].[dbo].[NC_EXP_IMP_HSDESC_table]</p>
            {html}
        </body>
      </html>
      """

    part1 = MIMEText(html, "html")
    message.attach(part1)


    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com",port, context = context) as server:
        server.login("tdmUsageAlert@gmail.com", password)
        server.sendmail(sender_email, recipients, message.as_string())

# Connect to SQL Server
conn = pyodbc.connect(
                        'Driver={SQL Server};'
                        'Server=SEVENFARMS_DB3;'
                        'Database=HSDESC;'
                        'UID=sa;'
                        'PWD=Harpua88;'
                        'Trusted_Connection=No;'
)


# Create Table
df = pd.read_sql_query('''
SELECT COUNT(DISTINCT [COMMODITY]) AS [# OF ROWS STARTING WITH DASHES], [CTY], [REPORTING_CTY]
FROM [HSDESC].[dbo].[NC_EXP_IMP_HSDESC_table] A
LEFT JOIN [Control].[dbo].[Data_Availability_Monthly] B
ON A.CTY = B.[DA_ISO_CODE2]
WHERE A.en LIKE '-%'
OR A.en LIKE ' -%'
GROUP BY [CTY], [REPORTING_CTY]
ORDER BY 1 DESC
			''', conn)

print(df)
if not df.empty:
    emailTable(df)
