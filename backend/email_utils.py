import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv
load_dotenv()


class EmailHandler:
    def __init__(self):
        # Gmail credentials stored securely in AWS Secrets Manager or .env
        self.email_user = "dipdap90@gmail.com"
        self.email_pass = "cnuk rgfb loih vyem"
        print(self.email_user)
        print(self.email_pass)
    def send_otp_email(self, recipient_email, otp):
        """Send OTP email via Gmail SMTP"""
        try:
            msg = MIMEMultipart()
            msg["From"] = self.email_user
            msg["To"] = recipient_email
            msg["Subject"] = "Your OTP Verification Code"

            body = f"""
            <html>
                <body>
                    <h3>Your OTP for verification is: <b>{otp}</b></h3>
                    <p>This OTP is valid for 10 minutes.</p>
                </body>
            </html>
            """

            msg.attach(MIMEText(body, "html"))

            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(self.email_user, self.email_pass)
                server.send_message(msg)

            print(f"OTP sent successfully to {recipient_email}")
            return True

        except Exception as e:
            print(f"Failed to send OTP: {str(e)}")
            return False
