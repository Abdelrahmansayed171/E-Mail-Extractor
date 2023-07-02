import imaplib, email, yaml

#Let us Open input authentication file
with open("auth.yml") as authFile:
    content = authFile.read()

#Load data from yaml file 
credentials = yaml.load(content, Loader=yaml.FullLoader)

#Load the userName and password from yaml file
user, password = credentials["user"], credentials["password"]

# set imap URL which is used in connection with GMAIL
imap_url = 'imap.gmail.com'

# Connection with GMail domain using SSL
myMail = imaplib.IMAP4_SSL(imap_url)

# Log in using your email, password
myMail.login(user, password)
print("logged into mailbox successfully!")

# Select the Inbox to fetch messages
myMail.select('Inbox', readonly=True)
print("Inbox Selected.")

try:
    # Load all Mail IDs in Inbox directory
    response_code, mail_ids = myMail.search(None, "ALL")
    # Extract IDs from mail_ids
    mail_id_list = mail_ids[0].decode().split()

    print("Response Code : {}".format(response_code))
    print("Mail IDs : {}\n".format(mail_id_list))

except Exception as e:
    print("INBOX - ErrorType : {}, Error : {}".format(type(e).__name__, e))


for id in mail_id_list:
    print("================== Start of Mail [{}] ====================".format(id))
    
    
    _, mail_data = myMail.fetch(id, '(RFC822)') ## Fetch mail data.

    message = email.message_from_bytes(mail_data[0][1]) ## Construct Message from mail data
    
    print("From       : {}".format(message.get("From")))
    print("To         : {}".format(message.get("To")))
    print("Bcc        : {}".format(message.get("Bcc")))
    print("Date       : {}".format(message.get("Date")))
    print("Subject    : {}".format(message.get("Subject")))

    print("Body : ")
    for part in message.walk():
        if part.get_content_type() == "text/plain":
            body_lines = part.as_string().split("\n")
            print("\n".join(body_lines[2:])) ### Print from the third line to the end
    

    print("================== End of Mail [{}] ====================\n".format(id))

myMail.close()