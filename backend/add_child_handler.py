import boto3
import json
import os
import random
import time
from datetime import datetime, timedelta
from backend.email_utils import EmailHandler
from botocore.exceptions import ClientError

class AddChildHandler:
    def __init__(self, region="ap-south-1"):
        self.dynamodb = boto3.resource(
            "dynamodb",
            aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
            aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
            region_name=region
        )
        self.parents_table = self.dynamodb.Table("Parents")
        self.children_table = self.dynamodb.Table("Students")
        self.email_handler = EmailHandler()

    def add_child_by_email(self, event):
        """Send OTP to child’s email for verification"""
        try:
            body = json.loads(event.get("body", "{}"))
            parent_email = body.get("parent_email")
            child_email = body.get("child_email")

            if not parent_email or not child_email:
                return self._response(400, {"error": "Both parent_email and child_email are required"})

            # Check if child exists in DynamoDB
            child_record = self.children_table.get_item(Key={"email": child_email}).get("Item")
            if not child_record:
                return self._response(404, {"error": "No child found with this email"})

            # Generate OTP
            otp = str(random.randint(100000, 999999))
            expiry = int(time.time()) + 600  # 10 minutes

            # Store OTP in child’s record
            self.children_table.update_item(
                Key={"email": child_email},
                UpdateExpression="SET pending_otp = :otp, otp_expiry = :exp",
                ExpressionAttributeValues={":otp": otp, ":exp": expiry}
            )

            # Send OTP email
            email_sent = self.email_handler.send_otp_email(child_email, otp)
            if not email_sent:
                return self._response(500, {"error": "Failed to send OTP email"})

            return self._response(200, {"message": "OTP sent successfully"})

        except ClientError as e:
            return self._response(400, {"error": str(e)})

    def _response(self, status, body):
        return {
            "statusCode": status,
            "headers": {"Content-Type": "application/json"},
            "body": json.dumps(body)
        }
