import pandas as pd
import smtplib
import os
import time

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Configuration

# Path to Excel file containing email data
EXCEL_PATH = r"C:/Users/Jhon/email_list.xlsx" # Replace with your Excel file path

# Base directory where attachments are stored
ATTACHMENT_BASE_PATH = r"C:/Users/Jhon/Downloads/attachments/" # Replace with your attachment directory

# SMTP (Outlook / Office365) settings 
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

# Sender credentials (use environment variables in production)
EMAIL = "your_email@abc.com" # Replace with your email
PASSWORD = "YOUR_APP_PASSWORD" # Replace with your app password
 
# CC recipient (same for all emails)
CC_EMAIL = "cc_email@abc.com" # Replace with CC email

# Load recipient data

df = pd.read_excel(EXCEL_PATH)

# Establish SMTP connection

server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
server.starttls()
server.login(EMAIL, PASSWORD)

# Send emails

for _, row in df.iterrows():

    receiver = row["Email"]
    name = row["EmailName"].strip()
    filename = row["Filepath"]

    attachment_path = os.path.join(ATTACHMENT_BASE_PATH, filename)

    # Create email message
    msg = MIMEMultipart()
    msg["From"] = EMAIL
    msg["To"] = receiver
    msg["Cc"] = CC_EMAIL
    msg["Subject"] = "Subject of the Email"

    greeting = f"Good Morning, {name},"

    # HTML email body
    body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6;">
        <p>{greeting}</p>

        <p> body of email goes here. This is a placeholder for the actual content of the email. </p>
    </body>
    </html>
    """

    msg.attach(MIMEText(body, "html"))

    # Attach file
    with open(attachment_path, "rb") as file:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(file.read())

    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename={filename}"
    )
    msg.attach(part)

    # Send email
    server.sendmail(
        EMAIL,
        [receiver, CC_EMAIL],
        msg.as_string()
    )

    print(f"Sent: {receiver} | Attachment: {filename}")

    # Throttle to avoid SMTP rate limits
    time.sleep(2)

# Cleanup

server.quit()
print("All emails sent successfully.")
