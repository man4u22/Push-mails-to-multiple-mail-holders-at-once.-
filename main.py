import requests
import base64
import pandas as pd


df = pd.read_excel(r"HR contacts_1.xlsx")


with open(r"Manoj_hp_resume1.pdf", "rb") as pdf_file:
    pdf_base64 = base64.b64encode(pdf_file.read()).decode("utf-8")


API_KEY = "SG.**********************************************"  
SENDGRID_URL = "https://api.sendgrid.com/v3/mail/send"
HEADERS = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}


batch = df.iloc[0:]

for index, row in  batch.iterrows():
    hr_name = str(row["Name"]).strip()
    hr_title = str(row["Title"]).strip()
    hr_email = str(row["Email"]).strip()
    hr_company = str(row["Company"]).strip()

    Subject = f"Application for Internship opportunity about 3-6 months"
    body_html = f"""
<html>
  <body style="font-family: Arial, sans-serif; line-height: 1.6; font-size: 14px; color: #333;">
    <p>Dear {hr_title},</p>

    <p>I hope this message finds you well. My name is Manoj H P and i am currently pursuing a <b>B.E
    Degree (7th sem)</b> on <b>Artificial Intelligence and Machine Learning</b> in <b>K S Institute of Technology</b>.
    I am reaching out to express my keen interest in internship opportunites at <b>{hr_company}</b>, as i 
    greatly look forward for the knowledge and experience that i would gain.</p>

    <p>During my academic journey, i have spent my time learning skills in the Machine Learning, Deep
    Learning, Data analysis and Web Development.further, i implemented my skills in building projects
    like:</p>
        <br><b>*Skin cancer detection using AI(Minor Project):</b></br>
        <br>Tech Stack Python: flask, Deep Learning, HTML, CSS, SQL, OpenCV, Tensorflow,PIL.</br>
        <br>"Built a skin cancer detection model using cnn,opencv and pillow for image processing, then deployed on a website where
        a user can check skin has cancerous cell or not and also book appointment through our website."</br>
        
        <br><b>*Portfolio Website:</b></br>
        <br>Tech Stack: HTML, CSS, Flask, SQL.</br>
        <br>"Built my portfolio website where i used HTML and css for the structure and design of webpage
        and utilized flask for connecting all the pages and made sure others can leave suugestion after
        seeing my website"</br>

        <p>I am eager to contribute my technical knowledge and problem-solving skills while gaining practical
        industry experience under your guidance. I would be grateful if you could consider my application for
        an internship opportunity at<b> {hr_company}</b>.</p>

        <p>I’ve attached my resume for your review. Please let me know if we could connect to discuss further.</p>
        <p><i>Roles</i>(<b>Intern</b>):Data Scientist,AI Product Engineer,AIML Enginner,Data Analyst,Software Enginner.</p>
        <p>Thank you for your time and consideration.</p>

        <br><b>Best regards,</b></br>
        <br>Manoj H P</br>
        <br><b>manojhp205@gmail.com</b></br>
    </body>
</html>    
    
    """

    payload = {
        "personalizations": [
            {
                "to": [{"email": "manojhp205@gmail.com", "name": hr_name}],
                "subject": Subject
            }
        ],
        "from": {
            "email": "1am22ai006@amceducation.in",  
            "name": "Manoj H P"
        },
        "content": [
            {"type": "text/html", "value": body_html}
        ],
        "attachments": [
            {
                "content": pdf_base64,
                "filename": "Resume.pdf",
                "type": "application/pdf",
                "disposition": "attachment"
            }
        ]
    }

    response = requests.post(SENDGRID_URL, headers=HEADERS, json=payload)

    if response.status_code in [200, 202]:
        print(f"Email sent to {hr_name} ({hr_email})")
    else:
        print(f"Failed for {hr_name} ({hr_email}): {response.text}")








    
    