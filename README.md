# Gmail Attachment Downloader

A Python tool to download attachments from starred Gmail emails and create an Excel report of email content.

## Problem Statement

I needed to process 200+ starred emails in a client's Gmail account, download all attachments, and create an organized Excel report containing email subjects and body content.

## Features

- Downloads attachments from all starred emails in Gmail
- Organizes attachments into folders based on email subject
- Creates an Excel spreadsheet with email subjects, body content, and attachment info
- Supports filtering for image attachments only
- Handles emails with multiple attachments or no attachments
- Customizable output locations

## Installation

1. Clone this repository:
```
git clone https://github.com/rajpathak-eng/gmail-attachment-downloader.git
cd gmail-attachment-downloader
```

2. Install required packages:
```
pip install -r requirements.txt
```

3. Set up Google API access:
   - Go to the [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project
   - Enable the Gmail API
   - Create OAuth 2.0 credentials (Desktop application)
   - Download the credentials file and rename it to `credentials.json`
   - Place it in the same directory as the script

## Usage

Basic usage:
```
python gmail_downloader.py
```

Customize output locations:
```
python gmail_downloader.py --attachment-dir="Download_Folder" --excel-file="Email_Report.xlsx"
```

Download only image attachments:
```
python gmail_downloader.py --image-only
```

Limit number of emails to process:
```
python gmail_downloader.py --max-emails=500
```

## How It Works

The script:
1. Authenticates with Gmail API
2. Retrieves all starred emails (using pagination to get all emails)
3. For each email:
   - Extracts subject, body content, and metadata
   - Downloads attachments to organized folders
4. Creates an Excel file with email data

## Requirements

- Python 3.6+
- Google API client libraries
- Pandas
- BeautifulSoup4

## Disclaimer

This tool is for personal use. Make sure you have proper authorization before accessing someone else's Gmail account.
