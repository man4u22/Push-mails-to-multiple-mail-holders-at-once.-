import pandas as pd
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os.path
import os
import mimetypes
import time
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# --- Gmail API Authentication Function ---
def gmail_auth():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret_563574727174-rfvq2lqvs17vc3livht54saa06teuhvt.apps.googleusercontent.com.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds

def send_email_with_attachment(service, sender_email, to_email, subject, body_html, file_path):
    message = MIMEMultipart()
    message['to'] = to_email
    message['from'] = sender_email
    message['subject'] = subject

    msg = MIMEText(body_html, 'html')
    message.attach(msg)

    # Attach the resume
    content_type, encoding = mimetypes.guess_type(file_path)
    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'

    main_type, sub_type = content_type.split('/', 1)
    with open(file_path, 'rb') as f:
        part = MIMEBase(main_type, sub_type)
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
        message.attach(part)
    
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    
    try:
        sent_message = service.users().messages().send(userId='me', body={'raw': raw_message}).execute()
        print(f'Email sent. Message ID: {sent_message["id"]}')
        return True
    except HttpError as error:
        print(f'An error occurred: {error}')
        return False


if __name__ == '__main__':
    try:
        df = pd.read_excel("HR contacts_1.xlsx")
    except FileNotFoundError:
        print("Error: 'HR contacts_1.xlsx' not found. Please check the file path.")
        exit()

    creds = gmail_auth()
    service = build('gmail', 'v1', credentials=creds)

    sender_email = "manojhp_aiml@ksit.edu.in"  
    resume_path = "Manoj_hp_resume1.pdf"
    

    if not os.path.exists(resume_path):
        print(f"Error: '{resume_path}' not found. Please check the file path.")
        exit()
    
    for index, row in df.iloc[2:70].iterrows():
        hr_name = str(row["Name"]).strip()
        hr_title = str(row["Title"]).strip()
        hr_email = str(row["Email"]).strip()
        hr_company = str(row["Company"]).strip()
        
        to_email = hr_email

        subject = f"Seeking an Internship about 3-6 months"
        body_html = f"""
        <html>
          <body style="font-family: Arial, sans-serif; line-height: 1.6; font-size: 14px; color: #333;">
            <p>Dear {hr_title},</p>
            <p>I hope this message finds you well. My name is Manoj H P and i am currently pursuing a <b>B.E Degree (7th sem)</b> on <b>Artificial Intelligence and Machine Learning</b> in <b>K S Institute of Technology</b>.
            I am reaching out to express my keen interest in internship opportunites at <b>{hr_company}</b>, as i greatly look forward for the knowledge and experience that i would gain.</p>
            <p>During my academic journey, i have spent my time learning skills in the Machine Learning, Deep Learning, Data analysis and Web Development.further, i implemented my skills in building projects like:</p>
            <b>*Skin cancer detection using AI(Minor Project):</b>
            <br>Tech Stack Python: flask, Deep Learning, HTML, CSS, SQL, OpenCV, Tensorflow,PIL.</br>
            <br>"Built a skin cancer detection model using cnn,opencv and pillow for image processing, then deployed on a website where a user can check skin has cancerous cell or not and also book appointment through our website."</br>
            
            <br><b>*Portfolio Website:</b></br>
            <br>Tech Stack: HTML, CSS, Flask, SQL.</br>
            <br>"Built my portfolio website where i used HTML and css for the structure and design of webpage and utilized flask for connecting all the pages and made sure others can leave suugestion after seeing my website"</br>
            <p>I am eager to contribute my technical knowledge and problem-solving skills while gaining practical industry experience under your guidance. I would be grateful if you could consider my application for an internship opportunity at<b> {hr_company}</b>.</p>
            <p>I’ve attached my resume for your review. Please let me know if we could connect to discuss further.</p>
            <p><i>Interested Roles</i>(<b>Intern</b>):Data Scientist,AI Product Engineer,AIML Enginner,Data Analyst,Software Enginner.</p>
            <p>Thank you for your time and consideration.</p>
            <br><b>Best regards,</b></br>
            <br>Manoj H P</br>
            <br><b>manojhp205@gmail.com</b></br>
          </body>
        </html>
        """
        
        print(f"Attempting to send email to {to_email}...")
        success = send_email_with_attachment(service, sender_email, to_email, subject, body_html, resume_path)
        
        if success:
            print(f"Email sent successfully to {to_email}")
        else:
            print(f"Failed to send email to {to_email}. See the error message above.")
        time.sleep(2)
        print(f"Waiting 2 seconds before sending the next email...")
       