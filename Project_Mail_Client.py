#MailClient is a Python program that leverages the Gmail API to enable users to send and receive emails through their Gmail accounts. 
#The application provides a graphical user interface (GUI) built using the wxPython library, allowing users to interact with their emails seamlessly.
#Created by Maria Kalendarova and Kaloian Piperkov


import os       #provides access to os-specific functionality
import json     # handling of JSON data
import wx       #gui    
from google.oauth2.credentials import Credentials       #handling OAuth 2.0 credentials for google APIs
from google_auth_oauthlib.flow import InstalledAppFlow  #implementing the OAuth 2.0 aouthorization flow
from googleapiclient.discovery import build             #builing a google API service
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError
from email import message_from_string 
from email.header import decode_header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email.message import EmailMessage
from email import encoders
import base64                   #encoding and decoding base64 data
import webbrowser               #opening url's in the user's browser
from bs4 import BeautifulSoup   #html parsing library for extracting content from email bodies

#viewing content of emails
class EmailViewer(wx.Frame):
    def __init__(self, parent, title, email_content):
        super(EmailViewer, self).__init__(parent, title=title, size=(600, 400))

        self.panel = wx.Panel(self)
        self.email_content = wx.TextCtrl(self.panel, style=wx.TE_MULTILINE | wx.TE_READONLY)
        self.email_content.SetValue(email_content)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.email_content, 1, wx.EXPAND | wx.ALL, 5)

        self.panel.SetSizer(sizer)
        self.Centre()
        self.Show()

#viewing the inbox emails, refreshing emails, option to send emails
class InboxWindow(wx.Frame):
    def __init__(self, parent, title, size, service, user_email, flow):
        super(InboxWindow, self).__init__(parent, title=title, size=size)
        self.panel = wx.Panel(self)
        self.service = service
        self.user_email = user_email
        
        self.inbox_label = wx.StaticText(self.panel, label="Inbox")
        self.email_list_ctrl = wx.ListCtrl(self.panel, style=wx.LC_REPORT | wx.BORDER_THEME)

        self.email_list_ctrl.InsertColumn(0, "Subject", width=300)
        self.email_list_ctrl.InsertColumn(1, "From", width=300)
        self.email_list_ctrl.InsertColumn(2, "Date", width=300)

        self.send_email_button = wx.Button(self.panel, label="Send Email")
        self.send_email_button.Bind(wx.EVT_BUTTON, self.on_send_email)
        self.refresh_button = wx.Button(self.panel, label="Refresh")
        self.refresh_button.Bind(wx.EVT_BUTTON, self.on_refresh)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.inbox_label, 0, wx.ALL, 5)
        sizer.Add(self.email_list_ctrl, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(self.send_email_button, 0, wx.ALL, 5)
        sizer.Add(self.refresh_button, 0, wx.ALL, 5)

        self.panel.SetSizer(sizer)

        #load emails into the list
        self.fetch_emails()
        self.flow = flow

        #set up timer for real-time updates
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.update_emails, self.timer)
        self.timer.Start(60000)  #1 minute

        #dictionary to store email data
        self.email_data_dict = {}  

    def on_refresh(self, event):
        self.fetch_emails()

    def fetch_emails(self):
        try:
            #clear existing items in the list control
            self.email_list_ctrl.DeleteAllItems()
            #clear the dictionary
            self.email_data_dict = {}

            results = self.service.users().messages().list(userId='me', labelIds=['INBOX']).execute()
            messages = results.get('messages', [])

            for idx, msg in enumerate(messages):
                message = self.service.users().messages().get(userId='me', id=msg['id']).execute()
                payload = message['payload']
                headers = payload['headers']

                subject = next((header['value'] for header in headers if header['name'] == 'Subject'), 'No Subject')
                sender = next((header['value'] for header in headers if header['name'] == 'From'), 'Unknown Sender')
                date = next((header['value'] for header in headers if header['name'] == 'Date'), 'No Date')

                index = self.email_list_ctrl.InsertItem(self.email_list_ctrl.GetItemCount(), subject)
                self.email_list_ctrl.SetItem(index, 1, sender)
                self.email_list_ctrl.SetItem(index, 2, date)

                # store the email ID in the dictionary with the index as the key
                self.email_data_dict[index] = msg['id']

        except Exception as e:
            wx.MessageBox(f"Error fetching emails: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def on_email_selected(self, event):
        #get the selected item index
        selected_index = event.GetIndex()
        #get the email ID using the index from the dictionary
        email_id = self.email_data_dict.get(selected_index, '')
        #fetch the full content of the selected email
        try:
            if email_id:
                selected_email = self.service.users().messages().get(userId='me', id=email_id).execute()
                email_body = self.get_email_body(selected_email['payload'])
                EmailViewer(None, title='Email Viewer', email_content=email_body)

        except Exception as e:
            wx.MessageBox(f"Error fetching email content: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

        self.fetch_emails()
        
    def update_emails(self, event):
        # update emails periodically
        self.fetch_emails()

    def on_send_email(self, event):
        try:
            # fetch credentials directly from the flow object
            credentials = self.flow.credentials
            # open the MailClient window for composing an email
            MailClient(None, title='Mail Client', size=(600, 600), credentials=credentials, user_email=self.user_email)

        except Exception as e:
            wx.MessageBox(f"Error opening MailClient: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def get_email_body(self, payload):
        #extract and decode the email body from the payload
        parts = payload.get('parts', [])
        for part in parts:
            if 'body' in part and 'data' in part['body']:
                data = part['body']['data']
                data = data.replace("-", "+").replace("_", "/")
                decoded_data = base64.b64decode(data)
                soup = BeautifulSoup(decoded_data, "lxml")
                return soup.body.text
        return 'No Content'

    def getEmails(self):
        try:
            #load credentials from the json file
            creds = None
            credentials_file = 'credentials.json'

            if os.path.exists(credentials_file):
                with open(credentials_file, 'r') as creds_file:
                    creds_data = creds_file.read()
                    if creds_data:
                        creds_info = json.loads(creds_data)
                        creds = Credentials.from_authorized_user_info(creds_info)

            # if credentials are not available or are invalid, ask the user to log in.
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file('credentials.json')
                    creds = flow.run_local_server(port=0)

                # save the access token in the json file for the next run
                with open(credentials_file, 'w') as creds_file:
                    creds_file.write(creds.to_json())

            # connect to the Gmail API
            service = build('gmail', 'v1', credentials=creds)
            # request a list of all the messages
            result = service.users().messages().list(userId='me').execute()
            # messages is a list of dictionaries where each dictionary contains a message id.
            messages = result.get('messages')

            for msg in messages:
                #get the message from its id
                txt = service.users().messages().get(userId='me', id=msg['id']).execute()
                #use try-except to avoid any Errors
                try:
                    #get value of 'payload' from dictionary 'txt'
                    payload = txt['payload']
                    headers = payload['headers']

                    #look for Subject and Sender Email in the headers
                    subject = next((d['value'] for d in headers if d['name'] == 'Subject'), 'No Subject')
                    sender = next((d['value'] for d in headers if d['name'] == 'From'), 'Unknown Sender')

                    #the body of the message is in Encrypted format
                    #get the data and decode it with base 64 decoder
                    parts = payload.get('parts')[0]
                    data = parts['body']['data']
                    data = data.replace("-", "+").replace("_", "/")
                    decoded_data = base64.b64decode(data)

                    #Now, the data obtained is in lxml. So, we will parse it with BeautifulSoup library
                    soup = BeautifulSoup(decoded_data, "lxml")
                    body = soup.body() # Extract text content from the body

                    #printing the subject, sender's email, and message
                    print("Subject:", subject)
                    print("From:", sender)
                    print("Message:", body)
                    print('\n')

                except Exception as e:
                    wx.MessageBox(f"Error processing email: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

        except Exception as e:
            wx.MessageBox(f"Error fetching emails: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

class LoginDialog(wx.Dialog):
    def __init__(self, parent, title):
        super(LoginDialog, self).__init__(parent, title=title, size=(250, 300))

        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        self.email_label = wx.StaticText(panel, label="Email:")
        self.email_text = wx.TextCtrl(panel)

        self.password_label = wx.StaticText(panel, label="Password:")
        self.password_text = wx.TextCtrl(panel, style=wx.TE_PASSWORD)

        self.email_label.SetPosition((-1000, -1000))
        self.email_text.SetPosition((-1000, -1000))
        self.password_label.SetPosition((-1000, -1000))
        self.password_text.SetPosition((-1000, -1000))

        self.login_button = wx.Button(panel, label="Login")
        self.login_button.Bind(wx.EVT_BUTTON, self.on_login)
        self.login_button.SetMinSize((200, 80))

        login_button_font = wx.Font(wx.FontInfo(10))
        self.login_button.SetFont(login_button_font)

        self.create_account_button = wx.Button(panel, label="Create Account")
        self.create_account_button.Bind(wx.EVT_BUTTON, self.create_account)
        self.create_account_button.SetMinSize((200, 80))

        create_account_button_font = wx.Font(wx.FontInfo(10))
        self.create_account_button.SetFont(create_account_button_font)

        sizer.Add(self.login_button, 0, wx.ALL, 20)
        sizer.Add(self.create_account_button, 0, wx.ALL, 20)

        panel.SetSizer(sizer)
        self.Centre()
        self.flow = None

    def on_login(self, event):
        email = self.email_text.GetValue()
        password = self.password_text.GetValue()

        #check if token file exists
        token_file = 'token.json'
        if os.path.exists(token_file):
            try:
                #load stored credentials from file
                with open(token_file, 'r') as token:
                    credentials_data = token.read()
                    if not credentials_data:
                        raise ValueError("Empty credentials file")
                    self.credentials = Credentials.from_authorized_user_info(json.loads(credentials_data), ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.compose', 'https://www.googleapis.com/auth/gmail.send'])   
                # set up the Gmail API service
                self.service = build('gmail', 'v1', credentials=self.credentials)
                # close the login dialog
                self.EndModal(wx.ID_OK)
                return

            except ValueError as ve:
                wx.MessageBox(f"Error loading stored credentials: {str(ve)}", "Error", wx.OK | wx.ICON_ERROR)
            except Exception as e:
                wx.MessageBox(f"Error loading stored credentials: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

        # if the token file is empty or not present, authenticate and get Gmail service
        self.flow = self.authenticate(email, password)
        
        if self.flow:
            # save the credentials for future use
            with open(token_file, 'w') as token:
                token.write(self.flow.credentials.to_json())
            # set up the Gmail API service
            self.service = build('gmail', 'v1', credentials=self.flow.credentials)
            # close the login dialog
            self.EndModal(wx.ID_OK)

    def authenticate(self, email, password):
        try:
            #set up OAuth 2.0 credentials
            self.flow = InstalledAppFlow.from_client_secrets_file(
                'path\\to\\credentials.json',                
                scopes=['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.compose', 'https://www.googleapis.com/auth/gmail.send'])
            
            credentials = self.flow.run_local_server(port=0)

            if not credentials:
                raise ValueError("Authentication failed or credentials are None")
            
            # build the Gmail API service
            service = build('gmail', 'v1', credentials=credentials)
            return service

        except ValueError as ve:
            wx.MessageBox(f"Authentication error: {str(ve)}", "Error", wx.OK | wx.ICON_ERROR)
            return None
        except Exception as e:
            wx.MessageBox(f"Authentication error: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
            return None

    def create_account(self, event):
        # open the default web browser to Gmail account creation page
        webbrowser.open("https://accounts.google.com/signup")

class MailClient(wx.Frame):
    def __init__(self, parent, title, size, credentials, user_email):
        super(MailClient, self).__init__(parent, title=title, size=size)
        self.panel = wx.Panel(self)
        self.credentials = credentials
        self.user_email = user_email

        self.to_label = wx.StaticText(self.panel, label="To:")
        self.to_text = wx.TextCtrl(self.panel)

        self.subject_label = wx.StaticText(self.panel, label="Subject:")
        self.subject_text = wx.TextCtrl(self.panel)

        self.body_label = wx.StaticText(self.panel, label="Body:")
        self.body_text = wx.TextCtrl(self.panel, style=wx.TE_MULTILINE)

        self.attach_button = wx.Button(self.panel, label="Attach Files", size=(120, 30))
        self.attach_button.Bind(wx.EVT_BUTTON, self.on_attach)
    
        self.send_button = wx.Button(self.panel, label="Send", size=(70, 30))
        self.send_button.Bind(wx.EVT_BUTTON, self.on_send)

        self.inbox_button = wx.Button(self.panel, label="Inbox", size=(70, 30))
        self.inbox_button.Bind(wx.EVT_BUTTON, self.open_inbox)

        self.status_text = wx.TextCtrl(self.panel, style=wx.TE_MULTILINE | wx.TE_READONLY)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.to_label, 0, wx.ALL, 5)
        sizer.Add(self.to_text, 0, wx.EXPAND | wx.ALL, 5)
        sizer.Add(self.subject_label, 0, wx.ALL, 5)
        sizer.Add(self.subject_text, 0, wx.EXPAND | wx.ALL, 5)
        sizer.Add(self.body_label, 0, wx.ALL, 5)
        sizer.Add(self.body_text, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(self.attach_button, 0, wx.ALL, 5)
        sizer.Add(self.send_button, 0, wx.ALIGN_RIGHT | wx.ALL, 5)
        sizer.Add(self.inbox_button, 0, wx.ALIGN_RIGHT | wx.ALL, 5)
        sizer.Add(self.status_text, 1, wx.EXPAND|wx.ALL, 5)

        self.panel.SetSizer(sizer)
        self.Show()

        #list to store attached files
        self.attachments = []

    def on_attach(self, event):
        with wx.FileDialog(self, "Choose files to attach", wildcard="All files (*.*)|*.*", style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_OK:
                new_attachments = fileDialog.GetPaths()
                # append newly attached files to the existing list
                self.attachments.extend(new_attachments)
                # clear existing text and display all attached file names in the status text
                self.status_text.Clear()
                self.status_text.AppendText("Attached Files:\n")
                for attachment in self.attachments:
                    self.status_text.AppendText(f"{os.path.basename(attachment)}\n")
    
    def on_send(self, event):
        receiver_email = self.to_text.GetValue()

        msg = MIMEMultipart()
        msg['Subject'] = self.subject_text.GetValue()
        msg['From'] = self.user_email
        msg['To'] = receiver_email
        msg.attach(MIMEText(self.body_text.GetValue(), 'plain'))

        # attach files to the email
        for attachment in self.attachments:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(open(attachment, 'rb').read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(attachment)}"')
            msg.attach(part)

        try:
            # create an instance of the Gmail API service
            service = build('gmail', 'v1', credentials=self.credentials)
            
            # send the email using Gmail API
            message = {'raw': base64.urlsafe_b64encode(msg.as_bytes()).decode('utf-8')}
            service.users().messages().send(userId=self.user_email, body=message).execute()

            wx.MessageBox("Email sent successfully!", "Success", wx.OK | wx.ICON_INFORMATION)
        except HttpError as e:
            wx.MessageBox(f"An error occurred while sending the email: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
        except Exception as e:
            wx.MessageBox(f"An error occurred: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def open_inbox(self, event):
        # close the current MailClient window
        self.Close()

if __name__ == '__main__':
    app = wx.App()

    login_dialog = LoginDialog(None, title='Login')

    if login_dialog.ShowModal() == wx.ID_OK:
        # user clicked "Login" in the dialog
        sender_email = login_dialog.email_text.GetValue()
        sender_password = login_dialog.password_text.GetValue()
        # authenticate and get Gmail service
        service = login_dialog.authenticate(sender_email, sender_password)

        if service:
            # create and show InboxWindow with the service passed as an argument
            inbox_window = InboxWindow(None, title='Inbox', size=(900, 700), service=service, user_email=sender_email, flow=login_dialog.flow)
            inbox_window.Bind(wx.EVT_LIST_ITEM_ACTIVATED, inbox_window.on_email_selected)
            inbox_window.Show()

            login_dialog.Destroy()

    app.MainLoop()
