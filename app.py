import imaplib, email, yaml

#Let us Open input authentication file
with open("auth.yml") as authFile:
    content = authFile.read()

credentials = yaml.load()
