import imaplib
import email
import yaml
import pandas as pd
import os
from email.header import decode_header


def create_excel_file():
    # FILE_NAME = 'E-mails.xlsx'

    data = {
        'From': [],
        'To': [],
        'Subject': [],
        'Body': [],
        'Attachments':[] 
    }

    # Create Data Frame with the given data
    data_frame = pd.DataFrame(data)

    # Create an Excel writer using pandas
    writer = pd.ExcelWriter('E-mails.xlsx', engine='xlsxwriter')
    
    # Write the DataFrame to an Excel sheet
    data_frame.to_excel(writer, sheet_name='Sheet1', index=False)

    print("Excel File Created!")
    writer.close()

def import_mail(from_user, to, subject, body, attachments=""):

    # Read the existing Excel file into a DataFrame
    reader = pd.read_excel('E-mails.xlsx')

    # Create a new record as a dictionary
    new_mail = {
        'From': [from_user],
        'To': [to],
        'Subject': [subject],
        'Body': [body],
        'Attachments': [attachments]
    }

    new_mail = pd.DataFrame(new_mail)


    # Concatenate the existing DataFrame with the new record
    data_frame = pd.concat([reader, new_mail], ignore_index=True)

    # used engine='openpyxl' because append operation is not supported by xlsxwriter
    writer = pd.ExcelWriter('E-mails.xlsx', engine='openpyxl', mode='a', if_sheet_exists="overlay")

    # append new dataframe to the excel sheet
    data_frame.to_excel(writer, index=False, header=False, startrow=len(reader) + 1)
    writer.close()



def display_message(msg):

    attachments_folder="Mail" + str(mail_id)
    from_mail = msg.get("From")
    to_mail = msg.get("To")
    cc = msg.get("Bcc")
    date = msg.get("Date")
    subject = decode_subject(msg.get("Subject"))
    body = ""

    print("================== Start of Mail [{}] ====================".format(id))
    print("From       : {}".format(from_mail))
    print("To         : {}".format(to_mail))
    print("Bcc        : {}".format(cc))
    print("Date       : {}".format(date))
    print("Subject    : {}".format(subject))

    print("Body : ")
    for part in msg.walk():
        if part.get_content_type() == "text/plain":
            body_lines = part.get_payload(decode=True).decode("utf-8")
            print(body_lines)
            body = body + body_lines
    print("================== End of Mail [{}] ====================\n".format(id))


def get_folder_name(mail_id):
    folder_name = "Attachments/Mail" + str(mail_id)

    # Get the current directory
    current_directory = os.getcwd()

    # Create the path for the new folder by joining Current path to new folder name together
    folder_path = os.path.join(current_directory, folder_name)

    if os.path.isdir(folder_path):
        return folder_path
    else:
        os.makedirs(folder_path, exist_ok=True)
        return folder_path

def extract_attachments(msg, mail_id):
    
    file_path = "No attachment found."

    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue

        if part.get('Content-Disposition') is None:
            continue

        filename = part.get_filename()
        

        if filename:
            folder_path = get_folder_name(mail_id)
            file_path = os.path.join(folder_path, filename)
            with open(file_path, 'wb') as downloaded_file:
                downloaded_file.write(part.get_payload(decode=True))
            downloaded_file.close()
    
    if file_path == "No attachment found.":
        return "No attachment found."
    else:
        return "Mail" + str(mail_id)

def decode_subject(subject):
    decoded_subject = ""
    if subject:
        decoded_parts = decode_header(subject)
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                if encoding:
                    part = part.decode(encoding)
                else:
                    part = part.decode("utf-8", "ignore")
            elif isinstance(part, str):
                part = part.encode("utf-8", "ignore").decode("utf-8")
            decoded_subject += part
    return decoded_subject



# Let us open the authentication file
with open("auth.yml") as authFile:
    content = authFile.read()

# Load data from the YAML file
credentials = yaml.load(content, Loader=yaml.FullLoader)

# Load the username and password from the YAML file
user, password = credentials["user"], credentials["password"]

# Set the IMAP URL for GMAIL
imap_url = 'imap.gmail.com'

# Connect to GMail domain using SSL
myMail = imaplib.IMAP4_SSL(imap_url)

# Log in using your email and password
myMail.login(user, password)
print("Logged into mailbox successfully!")

# Select the Inbox to fetch messages
myMail.select('Inbox', readonly=True)
print("Inbox selected.")

try:
    # Load all mail IDs in the Inbox directory
    response_code, mail_ids = myMail.search(None, "ALL")
    # Extract IDs from mail_ids
    mail_id_list = mail_ids[0].decode().split()

    print("Response Code: {}".format(response_code))
    print("Mail IDs: {}\n".format(mail_id_list))

except Exception as e:
    print("INBOX - ErrorType: {}, Error: {}".format(type(e).__name__, e))


create_excel_file()
# import_mail("abdelrahmansayed171@gmail.com", "orcaabs@gmail.com", "Hello buddy", "مرحبا اوركاا", "Orca/abbas/henna" )

for id in mail_id_list: 
    _, mail_data = myMail.fetch(id, '(RFC822)')  # Fetch mail data.
    message = email.message_from_bytes(mail_data[0][1])  # Construct message from mail data
    display_message()
    extract_attachments(message, id)
myMail.close()
