import imaplib
import email
from email.header import decode_header
import webbrowser
import os

username="info@vpms.club"
password="!2VPM'sPolyTechn#c"

imap = imaplib.IMAP4_SSL("imappro.zoho.com")
imap.login(username, password)

status, messages = imap.select("INBOX")
print(int(messages[0]))

messages = int(messages[0])
N = 2

for i in range(messages, messages-N, -1):
    # fetch the email message by ID
    res, msg = imap.fetch(str(i), "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            # parse a bytes email into a message object
            msg = email.message_from_bytes(response[1])
            # decode the email subject
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                # if it's a bytes, decode to str
                subject = subject.decode()
            # email sender
            from_ = msg.get("From")
            print("Subject:", subject)
            print("From:", from_)
            with open(f"vpm_seminar_2020_registrations.csv", "w") as output_csvfile:
                if "Registration:" in subject:
                    name = from_.split("<mail@vpms.club>")[0].strip()
                    print(f"Name: {name}")
                    # if the email message is multipart
                    if not msg.is_multipart():
                        # extract content type of email
                        content_type = msg.get_content_type()
                        # get the email body
                        body = msg.get_payload(decode=True).decode()
                        if content_type == "text/plain":
                            (email_part,contact_part,institute_part,designation_part) = ("","","","")
                            for line in body.splitlines():
                                email = line.split("email *:")
                                contact = line.split("Contact Number *:")
                                institute = line.split("Name of the Institute / Organisation *:")
                                designation = line.split("Designation *:")
                                
                                if len(email) > 1:
                                    email_part = email[1].strip()
                                
                                if len(contact) > 1:
                                    contact_part = contact[1].strip()
                                
                                if len(institute) > 1:
                                    institute_part = institute[1].strip()
                                
                                if len(designation) > 1:
                                    designation_part = designation[1].strip()
                            #print(f"{name}:{email_part}:{contact_part}:{institute_part}:{designation_part}")
                            output_csvfile.writelines(f"{name}:{email_part}:{contact_part}:{institute_part}:{designation_part}")
            print("="*100)
imap.close()
imap.logout()