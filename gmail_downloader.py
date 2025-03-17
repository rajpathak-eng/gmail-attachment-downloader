import os
import base64
import pandas as pd
import argparse
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from bs4 import BeautifulSoup
import re
import json

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def get_gmail_service():
    """Get authenticated Gmail API service."""
    creds = None
    # The file token.json stores the user's access and refresh tokens
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_info(
            json.loads(open('token.json').read()), SCOPES)
    
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return build('gmail', 'v1', credentials=creds)

def get_starred_emails(service):
    """Get all starred emails with pagination support."""
    messages = []
    next_page_token = None
    
    while True:
        results = service.users().messages().list(
            userId='me', 
            q='is:starred', 
            pageToken=next_page_token,
            maxResults=500  # Maximum allowed by the API
        ).execute()
        
        batch_messages = results.get('messages', [])
        if not batch_messages:
            break
            
        messages.extend(batch_messages)
        
        # Check if there are more pages
        next_page_token = results.get('nextPageToken')
        if not next_page_token:
            break
            
        print(f"Fetched {len(messages)} messages so far...")
    
    if not messages:
        print('No starred messages found.')
    else:
        print(f"Total starred messages found: {len(messages)}")
        
    return messages

def extract_email_data(service, message_id):
    """Extract subject, body, and attachments from an email."""
    msg = service.users().messages().get(userId='me', id=message_id, format='full').execute()
    
    # Get email headers
    headers = msg['payload']['headers']
    subject = ""
    from_email = ""
    date = ""
    
    for header in headers:
        if header['name'].lower() == 'subject':
            subject = header['value']
        elif header['name'].lower() == 'from':
            from_email = header['value']
        elif header['name'].lower() == 'date':
            date = header['value']
    
    # Process email parts
    parts = []
    if 'parts' in msg['payload']:
        parts = msg['payload']['parts']
    else:
        parts = [msg['payload']]
        
    body_text = ""
    attachments = []
    
    def process_parts(part_list):
        nonlocal body_text, attachments
        
        for part in part_list:
            # If this part has subparts, process them
            if 'parts' in part:
                process_parts(part['parts'])
                
            # Process body content
            if 'body' in part and 'data' in part['body']:
                mime_type = part.get('mimeType', '')
                if mime_type.startswith('text/'):
                    try:
                        body_data = base64.urlsafe_b64decode(part['body']['data']).decode('utf-8', errors='replace')
                        
                        # If it's HTML, try to extract just the text
                        if mime_type == 'text/html':
                            soup = BeautifulSoup(body_data, 'html.parser')
                            body_text += soup.get_text(separator='\n')
                        else:
                            body_text += body_data
                    except Exception as e:
                        print(f"Error decoding body: {e}")
            
            # Check if this part is an attachment
            if part.get('filename') and part['body'].get('attachmentId'):
                attachment = {
                    'id': part['body'].get('attachmentId'),
                    'filename': part['filename'],
                    'mimeType': part.get('mimeType', 'application/octet-stream')
                }
                attachments.append(attachment)
    
    process_parts(parts)
    
    return {
        'subject': subject,
        'body': body_text,
        'attachments': attachments,
        'from': from_email,
        'date': date
    }

def download_attachment(service, message_id, attachment_id, filename, download_dir):
    """Download an attachment and save it to disk."""
    attachment = service.users().messages().attachments().get(
        userId='me', messageId=message_id, id=attachment_id).execute()
    
    data = base64.urlsafe_b64decode(attachment['data'])
    
    # Create directory if it doesn't exist
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
    
    # Generate a safe filename
    safe_filename = re.sub(r'[^\w\-_\. ]', '_', filename)
    file_path = os.path.join(download_dir, safe_filename)
    
    # Handle duplicate filenames
    counter = 1
    base_name, ext = os.path.splitext(file_path)
    while os.path.exists(file_path):
        file_path = f"{base_name}_{counter}{ext}"
        counter += 1
    
    with open(file_path, 'wb') as f:
        f.write(data)
    
    return file_path

def is_image_attachment(mime_type):
    """Check if attachment is an image."""
    return mime_type.startswith('image/')

def sanitize_folder_name(name, max_length=50):
    """Create a safe folder name from email subject."""
    if not name or name.strip() == "":
        return "no_subject"
    
    # Remove invalid characters and limit length
    safe_name = re.sub(r'[^\w\-_\. ]', '_', name)
    return safe_name[:max_length]

def main():
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description='Download Gmail attachments and create Excel report')
    parser.add_argument('--attachment-dir', default='email_attachments', 
                        help='Directory to save attachments (default: email_attachments)')
    parser.add_argument('--excel-file', default='email_data.xlsx', 
                        help='Path to save Excel file (default: email_data.xlsx)')
    parser.add_argument('--image-only', action='store_true', 
                        help='Download only image attachments')
    parser.add_argument('--max-emails', type=int, default=None,
                        help='Maximum number of emails to process (default: all)')
    args = parser.parse_args()
    
    # Create directories for downloads
    download_dir = args.attachment_dir
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
    
    # Get Gmail service
    service = get_gmail_service()
    
    # Get starred emails (with pagination)
    starred_emails = get_starred_emails(service)
    
    # Limit the number of emails if specified
    if args.max_emails and len(starred_emails) > args.max_emails:
        print(f"Limiting to {args.max_emails} emails as requested")
        starred_emails = starred_emails[:args.max_emails]
    
    # Prepare data for Excel
    email_data = []
    
    # Process each email
    for i, email_msg in enumerate(starred_emails):
        print(f"Processing email {i+1}/{len(starred_emails)}")
        message_id = email_msg['id']
        
        try:
            # Extract email data
            data = extract_email_data(service, message_id)
            
            # Create a separate folder for each email (based on subject)
            if not data['subject']:
                email_folder_name = f"email_{i+1}"
            else:
                email_folder_name = sanitize_folder_name(data['subject'])
                
            # Add message ID to ensure uniqueness
            email_folder_name = f"{email_folder_name}_{message_id[-8:]}"
            email_folder = os.path.join(download_dir, email_folder_name)
            
            # Download attachments
            attachment_paths = []
            image_attachments = []
            other_attachments = []
            
            if data['attachments']:
                for attachment in data['attachments']:
                    if attachment['id']:  # Some attachments might not have an ID
                        try:
                            # Filter by image type if specified
                            if args.image_only and not is_image_attachment(attachment['mimeType']):
                                continue
                                
                            file_path = download_attachment(
                                service, message_id, attachment['id'], 
                                attachment['filename'], email_folder
                            )
                            
                            # Categorize attachments
                            if is_image_attachment(attachment['mimeType']):
                                image_attachments.append(file_path)
                            else:
                                other_attachments.append(file_path)
                                
                            attachment_paths.append(file_path)
                        except Exception as e:
                            print(f"Error downloading attachment {attachment['filename']}: {e}")
            
            # Add to email data list
            email_data.append({
                'Subject': data['subject'],
                'From': data['from'],
                'Date': data['date'],
                'Body': data['body'],
                'Total Attachments': len(attachment_paths),
                'Image Attachments': len(image_attachments),
                'Other Attachments': len(other_attachments),
                'Attachment Paths': '; '.join(attachment_paths) if attachment_paths else "No attachments"
            })
        except Exception as e:
            print(f"Error processing email {i+1}: {e}")
            # Add error entry to keep track of failed emails
            email_data.append({
                'Subject': f"Error processing email ID: {message_id}",
                'From': "Error",
                'Date': "Error",
                'Body': f"Error: {str(e)}",
                'Total Attachments': 0,
                'Image Attachments': 0,
                'Other Attachments': 0,
                'Attachment Paths': "Error"
            })
    
    # Create Excel file
    df = pd.DataFrame(email_data)
    
    # Make sure the directory for the Excel file exists
    excel_dir = os.path.dirname(args.excel_file)
    if excel_dir and not os.path.exists(excel_dir):
        os.makedirs(excel_dir)
    
    df.to_excel(args.excel_file, index=False)
    
    print(f"Process complete. {len(starred_emails)} emails processed.")
    print(f"Excel file created at: {args.excel_file}")
    print(f"Attachments downloaded to: {download_dir}")

if __name__ == '__main__':
    main()
