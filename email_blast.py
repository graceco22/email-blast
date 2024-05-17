import win32com.client as win32
import csv
from config import SENT_ON_BEHALF_OF_EMAIL

def read_mailing_list(filename):
    mailing_list = []
    with open(filename, 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            mailing_list.append((row[0], row[1]))   # first name column 1, email column 2
    return mailing_list

def read_email_body(filename):
    with open(filename, 'r') as file:
        return file.read()

def send_emails_with_confirmation(mailing_list, email_body):
    outlook = win32.Dispatch('outlook.application')
    
    check = input("Are you sure you want to send this email blast?\nRespond with Y/N\n")
    confirmation = check.upper() == "Y"

    if confirmation:
        for recipient_info in mailing_list:
            recipient_name, recipient_email = recipient_info

            mail = outlook.CreateItem(0)
            mail.SentOnBehalfOfName = SENT_ON_BEHALF_OF_EMAIL
            mail.To = recipient_email
            mail.Subject = "Hello, " + recipient_name.strip() + "!"

            customized_body = email_body.replace('<recipient_name>', recipient_name.strip())
            mail.HTMLBody = customized_body

            try:
                mail.Send()
                print(f"Email sent to {recipient_email}")
            except Exception as e:
                print(f"Failed to send email to {recipient_email}: {e}")
        
        print("All emails sent.")
    else:
        print("Email blast cancelled.")

mailing_list = read_mailing_list('mailing_list.csv')
email_body = read_email_body('email_body.html')

send_emails_with_confirmation(mailing_list, email_body)
