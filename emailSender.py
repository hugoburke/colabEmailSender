# Import necessary libraries
import pandas as pd
import requests
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Load your CSV with columns: Email, Name, Description
file_path = 'path_to_your_csv_file.csv'  # Path to your Companies CSV file
df = pd.read_csv(file_path)

# Email and API settings
from_email = 'your_email@example.com'  # Replace with your email address
app_password = 'your_email_app_password'  # Replace with your email app-specific password
cv_path = 'path_to_your_cv.pdf'  # Path to your CV file

# OpenAI API settings
CHATGPT_API_URL = "https://api.openai.com/v1/chat/completions"
CHATGPT_API_KEY = "your_openai_api_key"  # Replace with your OpenAI API key

# Function to create a personalized email message using ChatGPT
def personalize_message_with_chatgpt(company_name, description):
    """
    Generates a personalized email message using OpenAI GPT API.
    """
    prompt = (
        f"Write a personalized job-seeking email introducing me as a Software Engineer. "
        f"Express interest in {company_name} and mention relevant details based on the following description: {description}. "
        f"Ensure the email is confident, professional, and engaging. Format the response as follows:\n\n"
        f"Subject: <Your Subject Line>\n\n"
        f"Body:\n"
        f"<Your Email Body>"
    )

    headers = {
        "Authorization": f"Bearer {CHATGPT_API_KEY}",
        "Content-Type": "application/json"
    }
    data = {
        "model": "gpt-4",
        "messages": [
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": prompt}
        ]
    }

    response = requests.post(CHATGPT_API_URL, headers=headers, json=data)
    response_data = response.json()

    if response.status_code == 200 and 'choices' in response_data:
        content = response_data['choices'][0]['message']['content']
        # Extract Subject and Body
        subject_start = content.find("Subject:") + len("Subject:")
        body_start = content.find("Body:") + len("Body:")
        subject = content[subject_start:body_start].strip()
        body = content[body_start:].strip()
        return subject, body
    else:
        print(f"Error in ChatGPT API response: {response_data}")
        return None, None

# Function to sanitize the subject line
def sanitize_subject(subject):
    """Remove problematic characters or extra text from the subject line."""
    return subject.split("\n")[0].strip()

# Function to send email with personalized subject, body, and CV attachment
def send_email(to_email, company_name, description):
    """
    Sends an email with a personalized subject, body, and CV attachment.
    """
    subject, body = personalize_message_with_chatgpt(company_name, description)
    if not subject or not body:
        print(f"Skipping email to {to_email} due to error in generating email content.")
        return

    # Sanitize the subject line
    subject = sanitize_subject(subject)

    # Initialize the email message
    message = MIMEMultipart()
    message['From'] = from_email
    message['To'] = to_email
    message['Subject'] = subject
    message.attach(MIMEText(body, 'plain'))

    # Attach the CV
    with open(cv_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(cv_path)}')
        message.attach(part)

    # Send the email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(from_email, app_password)
        server.sendmail(from_email, to_email, message.as_string())
        print(f"Email sent to {company_name} ({to_email})")

# Loop through the dataframe and send emails
for index, row in df.iterrows():
    company_name = row['Name']
    description = row['Description']
    email = row['Email']

    # Skip rows where the email is empty or null
    if pd.isna(email) or email.strip() == "":
        print(f"Skipping row for {company_name} due to missing email.")
        continue

    # Send the email with the personalized message and CV attachment
    try:
        send_email(email, company_name, description)
        # Remove the row after a successful email send
        df = df.drop(index)
        print(f"Row for {company_name} ({email}) removed after successful email.")
    except Exception as e:
        print(f"Error sending email to {company_name} ({email}): {e}")

# Save the updated dataframe back to the CSV
df.to_csv(file_path, index=False)
print("Updated CSV saved.")
