import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

# Load the Excel file
df = pd.read_excel('Email.xlsx')

# Email configuration for Gmail
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_username = ''  # Replace with your Gmail address
smtp_password = ''  # Replace with your App Password

from_email = 'vitotest0102@gmail.com'
subject = 'Happy Birthday!'

# Function to send email
def send_email(to_email, subject, body):
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Upgrade the connection to a secure encrypted SSL/TLS connection
            server.login(smtp_username, smtp_password)
            server.sendmail(from_email, to_email, msg.as_string())
        print(f'Email sent to {to_email}')
    except Exception as e:
        print(f'Failed to send email to {to_email}. Error: {e}')

# Get today's date
today = datetime.today().strftime('%m-%d')

# Iterate through the DataFrame and send emails
for index, row in df.iterrows():
    birth_date = pd.to_datetime(row['Birth date']).strftime('%m-%d')
    if birth_date == today:
        first_name = row['First Name']
        last_name = row['Last Name']
        to_email = row['Email']
        body = f'''Dear {first_name} {last_name},
        
        Happy Birthday! {first_name}, Wish You All the best. May Allah give you everything what you want.
        
        Best wishes,
        Rogit Corp'''
        send_email(to_email, subject, body)

print('Birthday emails sent!')
