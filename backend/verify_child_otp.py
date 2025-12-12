import boto3
import json
import os
import time
from datetime import datetime
from botocore.exceptions import ClientError

class VerifyChildOTPHandler:
    def __init__(self, region="ap-south-1"):
        self.dynamodb = boto3.resource(
            "dynamodb",
            aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
            aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
            region_name=region
        )
        self.parents_table = self.dynamodb.Table("Parents")
        self.children_table = self.dynamodb.Table("Students")

    def verify_child_otp(self, event):
        """Verify OTP and link child to parent"""
        try:
            body = json.loads(event.get("body", "{}"))
            parent_email = body.get("parent_email")
            child_email = body.get("child_email")
            otp = body.get("otp")

            if not all([parent_email, child_email, otp]):
                return self._response(400, {"error": "Missing required fields"})

            # Get child record
            child = self.children_table.get_item(Key={"email": child_email}).get("Item")
            if not child:
                return self._response(404, {"error": "Child not found"})

            # Validate OTP
            if str(child.get("pending_otp")) != str(otp):
                return self._response(400, {"error": "Invalid OTP"})

            if time.time() > child.get("otp_expiry", 0):
                return self._response(400, {"error": "OTP expired"})

            # Update records: link child with parent
            self.children_table.update_item(
                Key={"email": child_email},
                UpdateExpression="REMOVE pending_otp, otp_expiry SET parent_email = :p",
                ExpressionAttributeValues={":p": parent_email}
            )

            self.parents_table.update_item(
                Key={"email": parent_email},
                UpdateExpression="ADD children :child",
                ExpressionAttributeValues={":child": {child_email}}
            )


            return self._response(200, {"message": "Child successfully linked to parent"})

        except ClientError as e:
            return self._response(400, {"error": str(e)})

    def _response(self, status, body):
        return {
            "statusCode": status,
            "headers": {"Content-Type": "application/json"},
            "body": json.dumps(body)
        }
