# Email Sender

A Python script for sending personalized emails to multiple recipients with optional PDF attachments.

## Features

- Send personalized emails to multiple recipients from an Excel file
- Support for Gmail, Yahoo Mail, and Outlook/Hotmail
- Optional PDF attachment functionality
- Personalized email content with recipient's name
- Customizable sender name and signature
- Secure email sending with TLS encryption
- User-friendly launcher with automatic dependency checking

## Requirements

- Python 3.7.1 (required for proper TLS encryption support)
- Required packages:
  - pandas
  - openpyxl

## Installation

1. Clone or download this repository
2. Install the required packages by running:
   ```
   python install.py
   ```
3. If you're using Anaconda, you may need to run:
   ```
   conda install python=3.7.1=h33f27b4_4
   ```

## Setup

1. Create an Excel file named `email-list.xlsx` with the following columns:
   - Email
   - Recipient First Name
   - Recipient Last Name

2. Create a text file named `subject.txt` containing your email subject line
   - Use `{name}` to insert the recipient's first name
   - Use `{fullName}` to insert the recipient's full name

3. Create a text file named `message.txt` containing your email message
   - Use `{name}` to insert the recipient's first name
   - Use `{fullName}` to insert the recipient's full name

4. (Optional) Create a text file named `sender_name.txt` containing your name
   - This will be used in the email signature and "From" field

5. (Optional) Place a PDF file named `document.pdf` in the project directory if you want to attach it to emails

## Usage

### Quick Start
Simply run:
```
python run.py
```
This will:
- Check for all required files
- Verify Python version compatibility
- Install missing dependencies
- Launch the email sender

### Manual Start
Alternatively, you can run the sender directly:
```
python sender.py
```

## Security Notes

- The script uses TLS encryption for secure email sending
- Your email password is only used for authentication and is not stored
- Make sure to use an app password if you have 2FA enabled on your email account

## File Structure

```
email-sender/
├── sender.py           # Main script
├── run.py             # Launcher script
├── install.py         # Installation script
├── email-list.xlsx    # Recipient information
├── subject.txt        # Email subject template
├── message.txt        # Email message template
├── sender_name.txt    # Sender's name
├── document.pdf       # Optional PDF attachment
└── README.md          # This file
```

## Troubleshooting

- If you encounter TLS encryption issues, make sure you're using Python 3.7.1
- For Gmail users, you might need to enable "Less secure app access" or use an App Password
- Make sure all required files are in the correct format and location
- Check that your email service's SMTP settings are correct
- Use `run.py` to automatically check for and fix common issues

## License

This project is open source and available under the MIT License.
