import imaplib
import email
import yaml
from email.header import decode_header

def display_message(msg):
    print("================== Start of Mail [{}] ====================".format(id))
    print("From       : {}".format(msg.get("From")))
    print("To         : {}".format(msg.get("To")))
    print("Bcc        : {}".format(msg.get("Bcc")))
    print("Date       : {}".format(msg.get("Date")))
    print("Subject    : {}".format(decode_subject(msg.get("Subject"))))

    print("Body : ")
    for part in msg.walk():
        if part.get_content_type() == "text/plain":
            body_lines = part.get_payload(decode=True).decode("utf-8")
            print(body_lines)
    print("================== End of Mail [{}] ====================\n".format(id))


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


for id in mail_id_list:
    if id == '6':
        _, mail_data = myMail.fetch(id, '(RFC822)')  # Fetch mail data.
        message = email.message_from_bytes(mail_data[0][1])  # Construct message from mail data
        display_message(message)


myMail.close()
