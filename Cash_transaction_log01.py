import gspread
from email.mime.text import MIMEText
from email.message import EmailMessage
import smtplib 
from datetime import datetime

gc = gspread.service_account("C:\\Users\\LENOVO\\Downloads\cash-transaction-log-f90e6b0a990c.json")

sh = gc.open_by_key("1qivrJdyKjeSJUUb0vnCs241K9ZM5fHeC_uT1HHy86c4")

wks = sh.worksheet("DEC'23")  

email_dictionary = {
    "Ahanf" : "ahnaftahmid012@gmail.com",
    "Bayazid" : "Bayazidhassan2003@gmail.com",
    "Ovi" : "asifaslamovi@gmail.com",
    "Kefaet" : "kefaet2003@gmail.com"}
#checking
# print(sh.sheet1.get('C8'))

EMAIL_ADDRESS = "uniqoxtech@gmail.com"      # Gmail setup
EMAIL_PASSWORD = "kgsyoknysytoqxqa"
host = "smtp.gmail.com"
port = 587

def send_email (recipient,name,type):
    if type=="normal":
        body = f'''Subject:message from Cash Transaction Log

Dear {name},

Your balance at Cash Transaction Log has been finished.
We encourage you to recharge your account at your earliest convenience.

Thank you.

Best Regards,
Cash Transaction Log
A UniqoXTech product
'''

    if type=="abnormal":
        body = f'''Subject:message from Cash Transaction Log 

Dear {name},

Your due at Cash Transaction Log has already been surpassed 500 BDT.
We are highly encouraging you to recharge your account at your earliest convenience.

Thank you.

Best Regards,
Cash Transaction Log
A UniqoXTech product
''' 

    if type=="wifi":
        body = f'''Subject:message from Cash Transaction Log 

Dear {name},

Your WIFI package is going to be finished just after 2 days.
To avoid any disruptions, we recommend you to renew the package.

Thank you.

Best Regards,
Cash Transaction Log
A UniqoXTech product
'''     
        
    smtp = smtplib.SMTP(host,port)

    status_code,response = smtp.ehlo()
    print(f"[*] Echoing the server: {status_code} {response}")
    status_code, response = smtp.starttls()
    print(f"[*] Starting TLS connection : {status_code} {response}")
    status_code, response = smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
    print(f"[*] Logging in : {status_code} {response}")
    smtp.sendmail(EMAIL_ADDRESS,recipient,body)
    print("Mail Sent Successfully!!")
    ##########

for row in range(23, 24):  # Adjust the range based on your needs
    for col in ["C", "D", "E", "F"]:
        cell_value = wks.acell(f"{col}{row}").value
        recipient = None
        name = None

        if cell_value.startswith("-") :
            print(cell_value)
            if col == "C":
                recipient = list(email_dictionary.values())[0]
                name = list(email_dictionary.keys())[0]
            elif col == "D":
                recipient = list(email_dictionary.values())[1]
                name = list(email_dictionary.keys())[1]
            elif col == "E":
                recipient = list(email_dictionary.values())[2]
                name = list(email_dictionary.keys())[1]
            else :
                recipient = list(email_dictionary.values())[3]
                name = list(email_dictionary.keys())[2] 

            if float(cell_value) > -500.00 :    
                send_email(recipient, name,"normal")       
            else:    # Additional condition: If the negative value is less than -500
                send_email(recipient, name,"abnormal")

current_date = datetime.now().day

if current_date == 8 :
    for name, recipient in email_dictionary.items():
        send_email(recipient, name, "wifi")

smtp.quit()   



    # message = EmailMessage()    
    # message = MIMEText(body)
    # message["Subject"] = subject
    # message["From"] = EMAIL_ADDRESS
    # message["To"] = recipient

    # with smtplib.SMTP("smtp.gmail.com", 587) as server:
    #     server.starttls()
    #     server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    #     server.sendmail(EMAIL_ADDRESS, recipient, message.as_string())