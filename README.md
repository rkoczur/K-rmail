# Mail Merge Sender

A simple Python/Tkinter application for sending mail merge emails.

## Features

- SMTP server configuration (host, port, username, password, TLS)
- Load Excel (`.xlsx`, `.xls`) or CSV files
- Select the email address field from loaded columns
- Edit text letter template and insert fields (in `{FieldName}` format)
- Optional SMTP connection types: None, TLS, SSL
- Preview and send

## Installation

1. Open a terminal in the project folder.
2. Install the dependencies:

```bash
python -m pip install -r requirements.txt
```

## Usage

```bash
python mail_merge.py
```

## Notes

- To load an Excel file, click the "Load Excel/CSV" button.
- Select the column containing the email addresses.
- In the letter template, insert field names in curly brackets, e.g., `{Name}` or `{Company}`.
- Before sending, verify the SMTP settings and sender email address.
