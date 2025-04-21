# NOTE: Some of the current Anaconda software does not support the encryption done by the starttls command.
# Therefore, you must use conda install python=3.7.1=h33f27b4_4 to downgrade from 3.7.3-h8c8aaf0_0 in order to rectify this problem.

# Import necessary modules and packages
import pandas as pd
from smtplib import SMTP, SMTPAuthenticationError, SMTPException
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import traceback
import sys
import time

def print_debug_info(sender_email):
    print("\n=== Debug Information ===")
    print(f"Sender Email: {sender_email}")
    print("SMTP Server: smtp.gmail.com")
    print("\nTroubleshooting Steps:")
    print("1. For Gmail users:")
    print("   a. Go to https://myaccount.google.com/lesssecureapps")
    print("   b. Turn ON 'Allow less secure apps'")
    print("   OR")
    print("   a. If you have 2FA enabled, go to https://myaccount.google.com/apppasswords")
    print("   b. Generate an App Password")
    print("   c. Use this App Password instead of your regular password")
    print("2. Check your spam/junk folder")
    print("3. Verify the recipient email addresses are correct")
    print("4. Make sure you're using Python 3.7.1")
    print("5. Check your internet connection")
    print("=======================\n")

def test_email_connection(smtp_server, email, password):
    try:
        print("\nTesting email connection...")
        server = SMTP(smtp_server, 587)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(email, password)
        print("✅ Connection test successful!")
        server.quit()
        return True
    except SMTPAuthenticationError:
        print("\n❌ Authentication failed. Please check your email and password.")
        return False
    except Exception as e:
        print(f"\n❌ Connection test failed: {str(e)}")
        return False

# Set Gmail as default SMTP server
serverStr = "smtp.gmail.com"

# Read the Excel file (NOTE: The information given is not real)
try:
    emailList = pd.read_excel("email-list.xlsx")
    print("✅ Successfully read email-list.xlsx")
except Exception as e:
    print(f"\n❌ Error reading email-list.xlsx: {str(e)}")
    print("Make sure the file exists and is in the correct format.")
    sys.exit(1)

# Grab all data from the Excel file
try:
    emails = emailList["Email"]
    firstNames = emailList["Recipient First Name"]
    lastNames = emailList["Recipient Last Name"]
    print(f"✅ Found {len(emails)} recipients in the Excel file")
except KeyError as e:
    print(f"\n❌ Error: Missing required column in email-list.xlsx: {str(e)}")
    print("Required columns: Email, Recipient First Name, Recipient Last Name")
    sys.exit(1)

# Get login information
print("\n=== Email Configuration ===")
senderEmail = input("What email address should the messages be sent from? ")
senderPassword = input("What is the password for the above email address? ")

# Test the connection first
if not test_email_connection(serverStr, senderEmail, senderPassword):
    print_debug_info(senderEmail)
    sys.exit(1)

# Configure SMTP
try:
    s = SMTP(serverStr, 587)
    s.ehlo()
    s.starttls()
    s.ehlo()
    s.login(senderEmail, senderPassword)
    print("✅ Successfully connected to Gmail SMTP server")
except SMTPAuthenticationError:
    print("\n❌ Authentication failed. Please check:")
    print("1. Your email and password are correct")
    print("2. Enable 'Less secure app access' at https://myaccount.google.com/lesssecureapps")
    print("3. If you have 2FA enabled, use an App Password instead")
    print_debug_info(senderEmail)
    sys.exit(1)
except Exception as e:
    print(f"\n❌ Connection error: {str(e)}")
    print_debug_info(senderEmail)
    sys.exit(1)

# Read the subject to be sent to all email addresses in the Excel file
try:
    with open("subject.txt", "r") as sbjFile:
        sbj = sbjFile.read()
    print("✅ Successfully read subject.txt")
except Exception as e:
    print(f"\n❌ Error reading subject.txt: {str(e)}")
    sys.exit(1)

# Read the message to be sent to all email addresses in the Excel file
try:
    with open("message.txt", "r", encoding="utf-8") as msgFile:
        msg = msgFile.read()
    print("✅ Successfully read message.txt")
except Exception as e:
    print(f"\n❌ Error reading message.txt: {str(e)}")
    sys.exit(1)

# Read sender name from file
try:
    with open("sender_name.txt", "r") as name_file:
        sender_name = name_file.read().strip()
    print("✅ Successfully read sender_name.txt")
except FileNotFoundError:
    sender_name = ""
    print("ℹ️ No sender_name.txt found - using email address only")

# Check if document.pdf exists and ask if user wants to attach it
attach_pdf = False
if os.path.exists("document.pdf"):
    while True:
        attach_choice = input("Would you like to attach document.pdf to the emails? (y/n): ").lower()
        if attach_choice in ['y', 'yes']:
            attach_pdf = True
            print("✅ PDF attachment enabled")
            break
        elif attach_choice in ['n', 'no']:
            break
        else:
            print("Please enter 'y' or 'n'")

# Begin sending messages
success_count = 0
fail_count = 0

print("\n=== Starting Email Sending ===")
for i in range(len(emails)):
    try:
        print(f"\nProcessing email {i+1} of {len(emails)}...")
        # Create a multipart message
        message = MIMEMultipart()
        message['From'] = f"{sender_name} <{senderEmail}>" if sender_name else senderEmail
        message['To'] = emails[i]
        
        # Edit the subject with information from the Excel file
        sbjWithFullName = sbj.replace("fullName", "%s %s" % (firstNames[i], lastNames[i]))
        sbjFull = sbjWithFullName.replace("name", firstNames[i])
        message['Subject'] = sbjFull

        # Edit the message with information from the Excel file
        msgWithFullName = msg.replace("fullName", "%s %s" % (firstNames[i], lastNames[i]))
        msgFull = msgWithFullName.replace("name", firstNames[i])
        
        # Attach the message body
        message.attach(MIMEText(msgFull, 'plain'))
        
        # Attach PDF if requested and exists
        if attach_pdf:
            with open("document.pdf", "rb") as pdf_file:
                pdf_attachment = MIMEApplication(pdf_file.read(), _subtype="pdf")
                pdf_attachment.add_header('Content-Disposition', 'attachment', filename="document.pdf")
                message.attach(pdf_attachment)

        # Send the email
        s.send_message(message)
        print(f"✅ Sent email to {firstNames[i]} {lastNames[i]} at {emails[i]}")
        success_count += 1
        
        # Add a small delay between emails to avoid rate limiting
        time.sleep(1)
        
    except Exception as e:
        print(f"❌ Failed to send email to {firstNames[i]} {lastNames[i]} at {emails[i]}")
        print(f"Error: {str(e)}")
        fail_count += 1

# Print summary
print("\n=== Email Sending Summary ===")
print(f"Successfully sent: {success_count}")
print(f"Failed to send: {fail_count}")
if fail_count > 0:
    print("\nFor failed emails, please check:")
    print("1. Email addresses are correct")
    print("2. Your email service's daily sending limits")
    print("3. Network connectivity")
    print_debug_info(senderEmail)

# Terminate connection once email sending has completed
s.quit()
print("\n=== Process Complete ===")
print("Please check your spam/junk folder if you don't see the emails in your inbox.")
