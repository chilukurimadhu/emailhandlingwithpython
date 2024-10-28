import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def send_email_via_outlook_smtp(to_address, subject, body, attachment_path=None):
    try:
        # SMTP server configuration for Outlook
        smtp_server = 'smtp.office365.com'
        smtp_port = 587
        outlook_user = 'madhuchilukuri9793@outlook.com'  # Replace with your Outlook email
        outlook_password = 'OutMad@1993'  # Replace with your Outlook password

        # Set up the MIME
        msg = MIMEMultipart()
        msg['From'] = outlook_user
        msg['To'] = to_address
        msg['Subject'] = subject

        # Attach the body with the msg instance
        msg.attach(MIMEText(body, 'plain'))

        # Attach a file if provided
        if attachment_path:
            attachment = open(attachment_path, 'rb')
            part = MIMEBase('application', 'octet-stream')
            part.set_payload((attachment).read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {attachment_path}")
            msg.attach(part)

        # Create the SMTP session
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Enable security

        # Login to the server
        server.login(outlook_user, outlook_password)

        # Send the email
        server.sendmail(outlook_user, to_address, msg.as_string())
        server.quit()

        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {str(e)}")

# Example usage
send_email_via_outlook_smtp(
    to_address='madhuchilukuri9793@outlook.com',
    subject='Test Email from Python',
    body='Hello, this is a test email sent from Python using Outlook.',
    attachment_path=None##'C:\path\to\your\file.txt'  # Optional, remove or set to None if no attachment
)
