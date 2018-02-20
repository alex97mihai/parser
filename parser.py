import imaplib
import email


if (input("Use defaults? y/n")=='y'):
    # Login details
    ORG_EMAIL = "ENTER DEFAULT DOMAIN HERE" # e.g. @gmail.com
    ACCOUNT_EMAIL = "ENTER DEFAULT USERNAME HERE" # the part before @
    FROM_PWD = "ENTER DEFAULT PASSWORD HERE"
    FROM_EMAIL = ACCOUNT_EMAIL + ORG_EMAIL
else:
    FROM_EMAIL = input("E-mail: ")
    FROM_PWD = input("Password: ")

# E-mail settings
SMTP_SERVER = "imap.gmail.com"
SMTP_PORT = 993

# Check e-mail type and get payload
def get_mpart(mail):
    maintype = mail.get_content_maintype()
    if maintype == 'multipart':
        for part in mail.get_payload():
            # This includes mail body AND text file attachements
            if part.get_content_maintype() == 'text':
                return part.get_payload()
        # No text at all
        return ""
    elif maintype == 'text':
        return mail.get_payload()


def readmail():
    # Try to connect to e-mail. Throw exception if unsuccessful
    try:
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL, FROM_PWD)
    except imaplib.IMAP4.error:
        print("Failed")
    
    # Select inbox folder to parse
    folder = input("Please select folder: ")
    mail.select(folder)
    
    # Get e-mail IDs
    type, data = mail.search(None, 'ALL')
    mail_ids = data[0]
    id_list = mail_ids.split()
    first_email_id = int(id_list[0])
    latest_email_id = int(id_list[-1])
        
    # Ask user for keywords:
    print("Please enter your keywords: ")
    keywords = [key for key in input().split()]
    data_output = []                            # empty list to hold output table
    
    # Iterate through e-mails
    for i in range(latest_email_id, first_email_id, -1):
        typ, data = mail.fetch(str(i), '(RFC822)')
        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_string(response_part[1].decode('utf-8'))
                email_subject = msg['subject']  # get subject
                email_from = msg['from']        # get sender
                email_datetime = msg['date']    # get date and time
                body = get_mpart(msg)           # get body
                words = body.split()            # split the body in words and save as list
                length = len(words)             # get number of words
                list1 = []                      # empty list to hold output
                list1.append(email_datetime)    # add date and time to the list
                
                # For first n-1 keywords 
                for keyIndex in range(0, len(keywords)-1):
                    c = False                   # assume keyword not found
                    # Search keyword in text
                    for index in range(0, length):
                        if words[index] == keywords[keyIndex]:
                            ans = ""
                            for j in range(index+1, length):
                                if words[j] == keywords[keyIndex+1] or words[j] == "--":
                                    break       # if next keyword reached, break
                                ans = ans + words[j] + ' '
                            list1.append(ans)   # keyword found, add result to list
                            c = True
                    if c == False:
                        list1.append("")

                # Do the same for last keyword
                c = False
                for index in range(0, length):
                    if words[index] == keywords[len(keywords)-1]:
                        ans = ""
                        for j in range(index+1, length):
                            if words[j] == "--":
                                break
                            ans = ans + words[j] + ' '
                        list1.append(ans)
                        c = True
                if c == False:
                    list1.append("")

                data_output.append(list1)       # add lists together 

    return data_output
    
output = readmail()
filename = input("Name for file: ")
out = open(filename, 'w', encoding='utf-8')
for row in output:
    for column in row:
        out.write('%s;' % column)
    out.write('\n')
out.close()


