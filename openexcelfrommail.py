import imaplib
import subprocess
import email
import os
import datetime
from email.header import decode_header
import openpyxl
import webbrowser

# Gmail login credentials
username = 'atthasedaanthony@gmail.com'
app_password = 'ebiwzittfrpyssyj'

# Folder to save attachments
download_folder = '/Users/yihao/Downloads'

# Sender's email to check for
specific_sender = 'atthaseda@gmail.com'

# Function to decode email subject and from field
def decode_mime_words(s):
    decoded_string, encoding = decode_header(s)[0]
    if isinstance(decoded_string, bytes):
        decoded_string = decoded_string.decode(encoding or 'utf-8')
    return decoded_string

# Connect to Gmail IMAP server
def connect_to_gmail():
    # Set up the connection
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    
    # Log in to the server
    mail.login(username, app_password)
    
    # Select the inbox
    mail.select("inbox")
    
    return mail

# Function to search and download attachments from emails after 11AM today
def download_and_open_attachment(mail):
    # Get today's date and time (after 5 PM)
    now = datetime.datetime.now()
    today = now.date()
    after_11am = datetime.datetime.combine(today, datetime.time(11, 0))  # 11:00 AM today
    after_11am_str = after_11am.strftime('%d-%b-%Y')

    # Search for emails from the specific sender received after 5 PM today
    status, messages = mail.search(None, f'(FROM "{specific_sender}" SINCE "{after_11am_str}")')
    
    # If there are any matching emails
    if status == "OK":
        for num in messages[0].split():
            # Fetch the email
            status, msg_data = mail.fetch(num, "(RFC822)")
            
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    # Parse the email content
                    msg = email.message_from_bytes(response_part[1])
                    
                    
                    # Check if the email has attachments
                    if msg.is_multipart():
                        for part in msg.walk():
                            # If part is an attachment
                            content_disposition = str(part.get("Content-Disposition"))
                            if "attachment" in content_disposition:
                                filename = part.get_filename()

         # Get the email subject and sender
                                subject = decode_mime_words(msg["Subject"])
                                from_ = decode_mime_words(msg.get("From"))
                                print(f"Subject: {subject}")
                                print(f"From: {from_}")
                                
                                if filename and filename.endswith('.xlsx'):
                                    print(f"Found Excel file: {filename}")
                                    
                                    # Download and save the file
                                    file_data = part.get_payload(decode=True)
                                    file_path = os.path.join(download_folder, filename)
                                    
                                    if not os.path.exists(download_folder):
                                        os.makedirs(download_folder)
                                    
                                    with open(file_path, 'wb') as f:
                                        f.write(file_data)
                                    print(f"File saved to {file_path}")

                                    # Open the Excel file with the default application
                                    open_excel_file(file_path)
            
                    else:
                        print("No attachments found.")
    else:
        print("No messages from the specified sender after 5 PM today.")

def open_excel_file(file_path):
    print(f"Opening {file_path}...")
    # Check the OS and use the appropriate command
    try:
        if os.name == 'nt':  # For Windows
            os.startfile(file_path)
        elif os.name == 'posix':  # For macOS or Linux
            # Use the 'open' command on macOS to open the file with the default app (Excel in this case)
            subprocess.run(['open', file_path])
        else:
            print("Unsupported OS for opening files.")
    except Exception as e:
        print(f"Error opening the file: {e}")

# Main function to execute the script
def main():
    mail = connect_to_gmail()
    download_and_open_attachment(mail)
    mail.logout()

if __name__ == "__main__":
    main()
