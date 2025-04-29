import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import openai

# Load your OpenAI API key
openai.api_key = 'your_openai_api_key'

# Function to generate email content using OpenAI
def generate_email_content():
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt="Write a professional email to inform about a new product launch.",
        max_tokens=150
    )
    return response.choices[0].text.strip()

# Read email addresses from Excel sheet
df = pd.read_excel('emails.xlsx')
email_list = df['Email'].tolist()

# Email details
sender_email = "your_email@example.com"
sender_password = "your_email_password"
subject = "New Product Launch"

# Create SMTP session
server = smtplib.SMTP('smtp.example.com', 587)
server.starttls()
server.login(sender_email, sender_password)

# Send email to each address
for recipient_email in email_list:
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    body = generate_email_content()
    msg.attach(MIMEText(body, 'plain'))

    server.sendmail(sender_email, recipient_email, msg.as_string())

# Close the SMTP session
server.quit()