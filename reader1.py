import imaplib
import email
from email.header import decode_header
import webbrowser
import os,sys
from time import sleep
import argparse
# os.path = "D:\traffic_images"
# account credentials
username = ""
password = ""

body=""
subject=""

def remove_html_tags(text):
    """Remove html tags from a string"""
    import re
    clean = re.compile('<.*?>')
    return re.sub(clean, '', text)

def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

# number of top emails to fetch
parser = argparse.ArgumentParser()
parser.add_argument("N", help="Number of emails to fetch",
                    type=int,default = 1)
args = parser.parse_args()
N = args.N

def read_mail():
    # create an IMAP4 class with SSL, use your email provider's IMAP server
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    # authenticate
    imap.login(username, password)

    # select a mailbox (in this case, the inbox mailbox)
    # use imap.list() to get the list of mailboxes
    status, messages = imap.select("INBOX")

    # total number of emails
    messages = int(messages[0])

    for i in range(messages, messages-N, -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                print("Subject:", subject)
                print("From:", From)
                # if the email message is multipart
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            print ("text/plain emails and skip attachments")
                            print(body)
                        elif "attachment" in content_disposition:
                            print("downloading attachment")
                            # download attachment
                            filename = part.get_filename()
                            if filename:
                                folder_name = clean(subject)
                                sub_folder_name = remove_html_tags(body)
                                #print(sub_folder_name)
                            if not os.path.isdir(folder_name):
                                # make a folder for this email (named after the subject)
                                os.mkdir(folder_name)
                            os.chdir(folder_name)

                            if not os.path.isdir(sub_folder_name):
                                os.mkdir(sub_folder_name)

                            filepath = os.path.join(sub_folder_name, filename)
                                # download attachment and save it
                            open(filepath, "wb").write(part.get_payload(decode=True))
                            os.chdir('../')
                            # os.chdir('../')

                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                         print ("only text email parts")
                         print(body)
                    if content_type == "text/html":
                    # if it's HTML, create a new HTML file and open it in browser
                        folder_name = clean(subject)
                        sub_folder_name = clean(body)
                    if not os.path.isdir(folder_name):
                        # make a folder for this email (named after the subject)
                        os.mkdir(folder_name)
                    os.chdir(folder_name)

                    if not os.path.isdir(sub_folder_name):
                        os.mkdir(sub_folder_name)
                    filename = "index.html"
                    filepath = os.path.join(sub_folder_name,filename)
                    # write the file
                    if not os.path.isdir(filepath):
                        open(filepath, "w").write(body)
                    # open in the default browser
                    webbrowser.open(filepath)
                #print("="*100)
    # close the connection and logout
    imap.close()
    imap.logout()

while True:
    read_mail()
    sleep(30)
