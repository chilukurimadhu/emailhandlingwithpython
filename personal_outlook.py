import requests
from msal import PublicClientApplication

def send_email_via_graph_api(to_address, subject, body):
    client_id = 'your_client_id'  # Your app's client ID
    tenant_id = 'consumers'  # Use 'consumers' for personal accounts
    scope = ["https://graph.microsoft.com/Mail.Send"]

    app = PublicClientApplication(client_id, authority=f"https://login.microsoftonline.com/{tenant_id}")

    # Interactive user login (opens a browser for sign-in)
    result = app.acquire_token_interactive(scopes=scope)

    if 'access_token' in result:
        token = result['access_token']
        graph_url = "https://graph.microsoft.com/v1.0/me/sendMail"
        email_data = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": body
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": to_address
                        }
                    }
                ]
            }
        }

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        response = requests.post(graph_url, json=email_data, headers=headers)

        if response.status_code == 202:
            print("Email sent successfully!")
        else:
            print(f"Failed to send email: {response.status_code}, {response.text}")

# Example usage
send_email_via_graph_api(
    to_address='recipient@example.com',
    subject='Test Email from Python',
    body='Hello, this is a test email sent from Python using Microsoft Graph API.'
)
