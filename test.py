import re
import imaplib
import email
import pandas as pd
import os
from email.header import decode_header
import re
import random
import string

cnt = 0

def create_excel_file():
    # FILE_NAME = 'E-mails.xlsx'

    data = {
        'From': [],
        'To': [],
        'CC': [],
        'Date': [],
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

def import_mail(from_user, to, cc, date, subject, body, attachments=""):

    # Read the existing Excel file into a DataFrame
    df = pd.read_excel('E-mails.xlsx')

    # Create a new record as a dictionary
    new_record = {
        'From': [from_user],
        'To': [to],
        'CC': [cc],
        'Date': [date],
        'Subject': [subject],
        'Body': [body],
        'Attachments': [attachments]
    }
    new_record = pd.DataFrame(new_record)

    # Append the new record to the DataFrame
    df = pd.concat([df, new_record], ignore_index=True)
    
    df.drop_duplicates()

    # Create an Excel writer using pandas
    writer = pd.ExcelWriter('E-mails.xlsx', engine='xlsxwriter')

    # Write the updated DataFrame to the Excel sheet
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Save the Excel file
    writer.close()

    print("New record appended to the Excel file!")



def display_message(msg, attachments, id):
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
            try:
                body_lines = part.get_payload(decode=True).decode("utf-8")
                print(body_lines)
                body += body_lines
            except UnicodeDecodeError:
                print("Skipping email [{}] due to decoding error.".format(id))
                cnt+=1
                return

    print("================== End of Mail [{}] ====================\n".format(id))

    import_mail(from_mail, to_mail, cc, date, subject, body, attachments)

def get_folder_name(mail_id):

    folder_name = "Attachments/Mail" + str(mail_id)

    # Get the current directory
    current_directory = os.getcwd()

    os.makedirs(os.path.join(current_directory,"Attachments"), exist_ok=True)


    # Create the path for the new folder by joining Current path to new folder name together
    folder_path = os.path.join(current_directory, folder_name)

    if os.path.isdir(folder_path):
        return folder_path
    else:
        os.makedirs(folder_path, exist_ok=True)
        return folder_path


def sanitize_filename(filename):
    try:
        # Decode the filename from UTF-8 and remove the encoding prefix
        decoded_filename = decode_header(filename)[0][0]
        if isinstance(decoded_filename, bytes):
            # If the filename is still bytes, decode it as UTF-8
            decoded_filename = decoded_filename.decode('utf-8')
        # Replace invalid characters with underscores
        sanitized_filename = re.sub(r'[\\/:*?"<>|]', '_', decoded_filename)
    except UnicodeDecodeError:
        # If decoding fails, generate a random filename
        random_string = ''.join(random.choices(string.ascii_letters + string.digits, k=8))
        return random_string

    sanitized_filename = re.sub(r'[\\/:*?"<>|]', '_', decoded_filename)
    return sanitized_filename



def extract_attachments(msg, mail_id):
    file_path = "No attachment found."

    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue

        if part.get('Content-Disposition') is None:
            continue

        filename = part.get_filename()

        if filename:
            try:
                folder_path = get_folder_name(mail_id)
                # Replace invalid characters in the filename
                filename = sanitize_filename(filename)
                file_path = os.path.join(folder_path, filename)
                with open(file_path, 'wb') as downloaded_file:
                    downloaded_file.write(part.get_payload(decode=True))
            except OSError as e:
                print("Skipping attachment due to error: {}".format(str(e)))
                cnt+= 1
                continue
            downloaded_file.close()

    if file_path == "No attachment found.":
        return file_path
    else:
        return "Mail" + str(mail_id)

def decode_subject(subject):
    decoded_subject = ""
    if subject:
        decoded_parts = decode_header(subject)
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                try:
                    if encoding:
                        part = part.decode(encoding)
                    else:
                        part = part.decode("utf-8", "ignore")
                except UnicodeDecodeError:
                    print("Skipping subject due to decoding error: {}".format(subject))
                    return ""
            elif isinstance(part, str):
                part = part.encode("utf-8", "ignore").decode("utf-8")
            decoded_subject += part
    return decoded_subject

def check_emails(mail):
    important = False

    important_emails=[
        "egypt-power",
        "morlock-motors",
        "gmail",
        "goldmaxpower"
        "yahoo",
        "outlook"
    ]

    for email in important_emails:
       if email in mail:
           important = True
    
    return important

def gui_handle(user, password):
    
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
        from_mail = message.get("From")
        if check_emails(from_mail):
            attachment = extract_attachments(message, id)
            display_message(message, attachment, id)
    
    myMail.close()
    print(cnt)
