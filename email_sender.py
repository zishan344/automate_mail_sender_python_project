# email_sender.py
import pandas as pd # type: ignore
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from config import EMAIL_ADDRESS, EMAIL_PASSWORD, SMTP_SERVER, SMTP_PORT

def send_email(to_address, subject, body):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_address
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        text = msg.as_string()
        server.sendmail(EMAIL_ADDRESS, to_address, text)
        server.quit()
        print(f"Email sent to {to_address}")
    except Exception as e:
        print(f"Failed to send email to {to_address}: {e}")

def read_recipients(file_path):
    df = pd.read_excel(file_path)
    return df

def main():
    recipients_file = 'recipients.xlsx'
    df = read_recipients(recipients_file)
    
    for index, row in df.iterrows():
        to_address = row['Email']
        name = row['Name']
        subject = f"Personalized Subject for {name}"
        body = f"Dear {name},\n\nThis is a personalized email.\n\nBest regards,\nMarouful Islam Zishan"
        send_email(to_address, subject, body)

if __name__ == "__main__":
    main()
