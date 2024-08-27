import tkinter as tk
from tkinter import filedialog, messagebox, Text
import pandas as pd
import mimetypes
import os
import base64
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def select_excel_file():
    filename = filedialog.askopenfilename(initialdir="/", title="Select an Excel File",
                                          filetypes=(("Excel files", "*.xlsx;*.xls"), ("all files", "*.*")))
    excel_file_entry.delete(0, tk.END)
    excel_file_entry.insert(0, filename)

def select_attachment_file():
    filename = filedialog.askopenfilename(initialdir="/", title="Select a File",
                                          filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))
    attachment_file_entry.delete(0, tk.END)
    attachment_file_entry.insert(0, filename)

def create_message_with_attachment(sender, to, subject, message_text, file):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    msg = MIMEText(message_text)
    message.attach(msg)

    content_type, encoding = mimetypes.guess_type(file)
    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    with open(file, 'rb') as fp:
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        encoders.encode_base64(msg)
    msg.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
    message.attach(msg)
    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}

def send_emails():
    subject = subject_entry.get()
    body = body_text.get("1.0", tk.END)
    attachment_path = attachment_file_entry.get()
    excel_path = excel_file_entry.get()

    if not os.path.exists(excel_path) or not os.path.exists(attachment_path):
        messagebox.showerror("Error", "Please specify both an Excel file and an attachment file.")
        return

    try:
        df = pd.read_excel(excel_path)
        service = service_gmail_api()

        for index, row in df.iterrows():
            receiver_email = row['emails']  # Adjust the column name as per your Excel sheet
            message = create_message_with_attachment('your_email@gmail.com', receiver_email, subject, body, attachment_path)
            send_message(service, 'me', message)
        
        messagebox.showinfo("Success", "All emails have been sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def service_gmail_api():
    scopes = ['https://www.googleapis.com/auth/gmail.compose']
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', scopes=scopes)
    creds = flow.run_local_server(port=0)
    return build('gmail', 'v1', credentials=creds)

def send_message(service, user_id, message):
    try:
        message = service.users().messages().send(userId=user_id, body=message).execute()
    except HttpError as error:
        print('An error occurred: %s' % error)

app = tk.Tk()
app.title('Email Sender')

frame = tk.Frame(app)
frame.pack(pady=20, padx=20)

tk.Label(frame, text="Subject:").grid(row=0, column=0, sticky='e')
subject_entry = tk.Entry(frame, width=40)
subject_entry.grid(row=0, column=1)

tk.Label(frame, text="Body:").grid(row=1, column=0, sticky='ne')
body_text = Text(frame, height=10, width=30)
body_text.grid(row=1, column=1)

tk.Label(frame, text="Excel File:").grid(row=2, column=0, sticky='e')
excel_file_entry = tk.Entry(frame, width=30)
excel_file_entry.grid(row=2, column=1)
tk.Button(frame, text="Browse...", command=select_excel_file).grid(row=2, column=2)

tk.Label(frame, text="Attachment File:").grid(row=3, column=0, sticky='e')
attachment_file_entry = tk.Entry(frame, width=30)
attachment_file_entry.grid(row=3, column=1)
tk.Button(frame, text="Browse...", command=select_attachment_file).grid(row=3, column=2)

send_button = tk.Button(frame, text="Send Emails", command=send_emails)
send_button.grid(row=4, columnspan=3, pady=10)

app.mainloop()
