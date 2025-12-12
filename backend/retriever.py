"""
Retriever Function for Parent Portal
------------------------------------
Retrieves student assignment and evaluation files from an S3 bucket
based on metadata (class_code, assignment_name, college_email, student_email, etc.)
and returns presigned URLs for each matching file.

✅ Works for:
    - Assignments
    - Evaluation Reports

✅ Uses AWS credentials from `.env` file
✅ Compatible with local and Lambda environments
"""

import os
import json
import boto3
import logging
from botocore.exceptions import ClientError
from dotenv import load_dotenv

# -------------------------------------------------
# Load environment variables from .env
# -------------------------------------------------
load_dotenv()

AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY")
AWS_REGION = os.getenv("AWS_REGION", "us-east-1")
BUCKET = os.getenv("S3_BUCKET_NAME")

if not AWS_ACCESS_KEY_ID or not AWS_SECRET_ACCESS_KEY:
    raise ValueError("❌ Missing AWS credentials in .env file")

# -------------------------------------------------
# Logging setup
# -------------------------------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# -------------------------------------------------
# Initialize S3 Client
# -------------------------------------------------
s3 = boto3.client(
    "s3",
    aws_access_key_id=AWS_ACCESS_KEY_ID,
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
    region_name=AWS_REGION
)

DEFAULT_EXPIRES = 3600  # 1 hour

# -------------------------------------------------
# Helper: Generate presigned URL
# -------------------------------------------------
def generate_presigned_url(bucket, key, expires_in=DEFAULT_EXPIRES):
    try:
        return s3.generate_presigned_url(
            "get_object",
            Params={"Bucket": bucket, "Key": key},
            ExpiresIn=expires_in
        )
    except Exception as e:
        logger.error(f"Error generating URL for {key}: {e}")
        return None

# -------------------------------------------------
# Helper: Match object metadata with filters
# -------------------------------------------------
def matches_metadata(metadata, filters):
    for k, v in filters.items():
        if not v:
            continue
        meta_val = metadata.get(k.lower()) or metadata.get(k) or ""
        if v.lower() not in meta_val.lower():
            return False
    return True

# -------------------------------------------------
# Core Retriever Class
# -------------------------------------------------
class ParentPortalRetriever:
    def __init__(self, bucket=BUCKET):
        self.bucket = bucket
        self.s3 = s3

    # Fetch metadata safely
    def get_metadata(self, key):
        try:
            head = self.s3.head_object(Bucket=self.bucket, Key=key)
            return {k.lower(): v for k, v in head.get("Metadata", {}).items()}
        except ClientError as e:
            logger.warning(f"Failed to fetch metadata for {key}: {e}")
            return {}

    # Retrieve presigned URLs by filtering metadata
    def retrieve_files(
        self, student_email="", class_code="", assignment_name="", college_email="", file_type="both", expires_in=DEFAULT_EXPIRES
    ):
        results = []
        filters = {
            "student_email": student_email,
            "class_code": class_code,
            "assignment_name": assignment_name,
            "college_email": college_email
        }

        try:
            paginator = self.s3.get_paginator("list_objects_v2")
            pages = paginator.paginate(Bucket=self.bucket)

            for page in pages:
                for obj in page.get("Contents", []):
                    key = obj["Key"]
                    metadata = self.get_metadata(key)

                    # Filter by type if needed
                    if file_type != "both":
                        obj_type = metadata.get("type", "")
                        if file_type.lower() not in obj_type.lower():
                            continue

                    if matches_metadata(metadata, filters):
                        url = generate_presigned_url(self.bucket, key, expires_in)
                        if url:
                            results.append({
                                "Key": key,
                                "Metadata": metadata,
                                "PresignedUrl": url
                            })
            return {"statusCode": 200, "results": results}

        except ClientError as e:
            if e.response["Error"]["Code"] == "AccessDenied":
                logger.warning("AccessDenied — switching to fallback mode.")
                return self.fallback_retrieve(filters, file_type, expires_in)
            else:
                logger.error(f"ClientError: {e}")
                return {"statusCode": 500, "error": str(e)}

    # Fallback if list_objects_v2 not allowed
    def fallback_retrieve(self, filters, file_type, expires_in):
        results = []
        possible_prefixes = []

        class_code = filters.get("class_code", "")
        student = filters.get("student_email", "")
        assignment = filters.get("assignment_name", "")

        if class_code and student and assignment:
            possible_prefixes.extend([
                f"submissions/{class_code}/{assignment}/{student}",
                f"evaluations/{class_code}/{assignment}/{student}"
            ])
        elif class_code:
            possible_prefixes.extend([
                f"submissions/{class_code}/",
                f"evaluations/{class_code}/"
            ])

        logger.info(f"Fallback prefixes: {possible_prefixes}")

        for prefix in possible_prefixes:
            try:
                metadata = self.get_metadata(prefix)
                if matches_metadata(metadata, filters):
                    url = generate_presigned_url(self.bucket, prefix, expires_in)
                    if url:
                        results.append({
                            "Key": prefix,
                            "Metadata": metadata,
                            "PresignedUrl": url
                        })
            except ClientError as e:
                if e.response["Error"]["Code"] == "404":
                    logger.info(f"No object found for {prefix}")
                else:
                    logger.warning(f"Fallback HeadObject failed for {prefix}: {e}")

        if results:
            return {"statusCode": 200, "results": results}
        else:
            return {"statusCode": 404, "message": "No objects found (fallback mode)."}


# -------------------------------------------------
# Lambda-style handler
# -------------------------------------------------
def handler(event, context=None):
    bucket = event.get("bucket") or BUCKET
    student_email = event.get("student_email", "")
    class_code = event.get("class_code", "")
    assignment_name = event.get("assignment_name", "")
    college_email = event.get("college_email", "")
    file_type = event.get("file_type", "both")
    expires_in = event.get("expires_in", DEFAULT_EXPIRES)

    retriever = ParentPortalRetriever(bucket)
    response = retriever.retrieve_files(
        student_email=student_email,
        class_code=class_code,
        assignment_name=assignment_name,
        college_email=college_email,
        file_type=file_type,
        expires_in=expires_in
    )
    return response


# -------------------------------------------------
# Local Test
# -------------------------------------------------
if __name__ == "__main__":
    test_event = {
        "bucket": "submitted-assignments",
        "student_email": "bilal@inikola.com",
        "class_code": "CLS-9F16E906",
        "assignment_name": "assignment1",
        "college_email": "example@college.com",
        "file_type": "both",
        "expires_in": 1200
    }
    print(json.dumps(handler(test_event), indent=2))
