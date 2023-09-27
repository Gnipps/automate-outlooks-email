import requests
import schedule
import time
from datetime import datetime
import pandas as pd

# Microsoft Graph API endpoint
graph_api_url = "https://graph.microsoft.com/v1.0/me/sendMail"

# Your Azure App Registration details
client_id = "YOUR_CLIENT_ID"
tenant_id = "YOUR_TENANT_ID"
client_secret = "YOUR_CLIENT_SECRET"

# Your email and recipient's email
sender_email = "your_email@example.com"
recipient_email = "recipient@example.com"

# Function to send an email with the specified content
def send_email():
    # Create a 3x4 table with your desired content
    table_data = {
        "Date": [datetime.now().strftime("%Y-%m-%d")],
        "A1": ["Content for A1"],
        "A2": ["Content for A2"],
    }

    df = pd.DataFrame(table_data)

    # Convert the DataFrame to HTML
    table_html = df.to_html(index=False)

    # Compose the email body
    email_body = f"<html><body><h2>Daily Report</h2>{table_html}</body></html>"

    # Request headers
    headers = {
        "Content-Type": "application/json",
    }

    # Request payload
    data = {
        "message": {
            "subject": "Daily Report",
            "body": {
                "contentType": "HTML",
                "content": email_body,
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": recipient_email,
                    }
                }
            ],
        },
        "saveToSentItems": "true",
    }

    # Get the access token
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "resource": "https://graph.microsoft.com",
    }
    token_response = requests.post(token_url, data=token_data)
    token = token_response.json().get("access_token")

    # Send the email
    response = requests.post(graph_api_url, headers=headers, json=data, headers={"Authorization": f"Bearer {token}"})

    if response.status_code == 202:
        print("Email sent successfully")
    else:
        print(f"Failed to send email. Status code: {response.status_code}")

# Schedule the job to run daily at 9 AM
schedule.every().day.at("09:00").do(send_email)

# Run the scheduler
while True:
    schedule.run_pending()
    time.sleep(1)
