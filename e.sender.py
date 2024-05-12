import pandas as pd
from datetime import datetime
import smtplib
from email.message import EmailMessage


def send_email(recipient, subject, msg):
    gmail_id = "gmail"
    gmail_paswword = "password"
    
    email = EmailMessage()
    email["Subject"] = subject
    email["From"] = gmail_id
    email["To"] = recipient
    email.set_content(msg)
    
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as gmail_obj :
        gmail_obj.ehlo()
        gmail_obj.login(gmail_id, gmail_paswword)
        gmail_obj.send_message(email)
    print("Email sent to " + str(recipient) + "with subject " + str(subject) + "and message: " + str(msg))

def send_bday_email(b_file):
    bdays_df = pd.read_excel(b_file)
    today = datetime.now().strftime("%m-%d")
    year_now = datetime.now().strftime("%Y")
    sent_index = []
    
    for idx, item in bdays_df.iterrows():
        bday = item['Birthday'].to_pydatetime().strftime('%m-%d')
        
        if ( today == bday ) and year_now not in str (item["Last Sent"]):
            msg= "happy Birthday" + str(item["Name"])
            send_email(item["Email"], "Happy Birthday", msg)
            sent_index.append(idx)
            
    for idx in  sent_index:
        bdays_df.loc(b_file, index= False)
        
if __name__ == "__main__":
    send_bday_email(bday_file ="file.xls")