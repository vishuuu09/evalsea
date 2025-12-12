# evaluator.py (UPDATED)
import json
import fitz  # PyMuPDF
import openai
import boto3
import io
import requests
import tiktoken
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.lib.enums import TA_CENTER
from openai import OpenAI
import os
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
import datetime
from botocore.exceptions import ClientError
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import base64
from docx import Document as DocxDocument
import mimetypes

# Optional OCR fallback
try:
    from PIL import Image
    import pytesseract
    PYTESSACT_AVAILABLE = True
except Exception:
    PYTESSACT_AVAILABLE = False

load_dotenv()

client = OpenAI(api_key=os.getenv("OPEN_AI_API_KEY"))


class Processor:
    def __init__(self):
        # S3 client used across methods
        self.S3_CLIENT = boto3.client(
            "s3",
            aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
            aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
            region_name=os.getenv("AWS_REGION", "us-east-1")
        )
        # DynamoDB resource and table references
        self.DDB = boto3.resource(
            "dynamodb",
            aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
            aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
            region_name=os.getenv("AWS_REGION", "us-east-1")
        )
        # Name of the students table
        self.STUDENTS_TABLE = self.DDB.Table(os.getenv("STUDENTS_TABLE_NAME", "Students"))
        # Name of class table (for retrieving department/college mail)
        self.CLASS_TABLE = self.DDB.Table(os.getenv("CLASS_TABLE_NAME", "Class"))

        self.S3_BUCKET_NAME = os.getenv("EVALUATED_REPORTS_BUCKET", "evaluated-reports")
        # bucket to store excel trackers
        self.EXCEL_BUCKET = os.getenv("ASSIGNMENT_EXCEL_BUCKET", "assignment-excels")
        # bucket where assignment master files are stored
        self.GENERATED_ASSIGNMENTS_BUCKET = os.getenv("GENERATED_ASSIGNMENTS_BUCKET", "generated-assignments")

        # Textract client for OCR (preferred for handwriting)
        try:
            self.TEXTRACT = boto3.client("textract",
                                        aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
                                        aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
                                        region_name=os.getenv("AWS_REGION", "us-east-1"))
        except Exception:
            self.TEXTRACT = None

    # ---------- PDF extraction ----------
    def read_pdf(self, file_bytes):
        text = ""
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for page in doc:
            text += page.get_text() + "\n"
        return text

    def extract_images_from_pdf(self, file_bytes):
        """
        Returns list of tuples: (img_name, img_bytes, ext)
        """
        images = []
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for page_index in range(len(doc)):
            page = doc[page_index]
            image_list = page.get_images(full=True)
            for img_index, img_meta in enumerate(image_list):
                xref = img_meta[0]
                base_image = doc.extract_image(xref)
                img_bytes = base_image["image"]
                ext = base_image.get("ext", "png")
                img_name = f"page{page_index+1}_img{img_index+1}.{ext}"
                images.append((img_name, img_bytes, ext))
        return images

    # ---------- docx extraction ----------
    def read_docx(self, file_bytes):
        """
        Extract text and images from a docx file (bytes).
        Returns text (string) and images list: (img_name, img_bytes, ext)
        """
        text_parts = []
        images = []
        try:
            # python-docx doesn't accept bytes directly; save to buffer and open
            buffer = io.BytesIO(file_bytes)
            doc = DocxDocument(buffer)
            for para in doc.paragraphs:
                text_parts.append(para.text)
            # extract images
            for rel in doc.part._rels:
                rel_obj = doc.part._rels[rel]
                if "image" in rel_obj.target_ref:
                    image_part = rel_obj.target_part
                    img_bytes = image_part.blob
                    # infer extension from content type if possible
                    ctype = image_part.content_type
                    ext = mimetypes.guess_extension(ctype) or ".png"
                    img_name = os.path.basename(rel_obj.target_ref)
                    images.append((img_name, img_bytes, ext.replace('.', '')))
        except Exception as e:
            print(f"Failed to read docx: {e}")
        return "\n".join(text_parts), images

    # ---------- ipynb extraction ----------
    def read_ipynb(self, file_bytes):
        """
        Extract text from ipynb: markdown cells, code cells, outputs, and attachments.
        Returns combined textual content and images list (img_name, img_bytes, ext)
        """
        try:
            nb = json.loads(file_bytes.decode("utf-8"))
        except Exception as e:
            # fallback: sometimes file_bytes may already be str
            nb = json.loads(file_bytes)

        text_parts = []
        images = []
        cells = nb.get("cells", [])
        for i, cell in enumerate(cells):
            ctype = cell.get("cell_type", "")
            if ctype == "markdown":
                text_parts.append("".join(cell.get("source", [])))
            elif ctype == "code":
                # include code as part of evaluation
                code_text = "".join(cell.get("source", []))
                text_parts.append(f"### Code cell {i+1}:\n{code_text}")
                # outputs: textual outputs
                for out in cell.get("outputs", []):
                    if "text" in out:
                        text_parts.append("Output:\n" + "".join(out.get("text", [])))
                    # images in outputs (data:image/png;base64)
                    data = out.get("data", {})
                    for k, v in data.items():
                        if k.startswith("image/"):
                            # v could be list or string
                            b64 = "".join(v) if isinstance(v, list) else v
                            try:
                                img_bytes = base64.b64decode(b64)
                                ext = k.split("/")[-1]
                                img_name = f"notebook_cell{i+1}_output.{ext}"
                                images.append((img_name, img_bytes, ext))
                            except Exception:
                                pass
            # attachments (in markdown cells)
            attachments = cell.get("attachments", {})
            for name, attachment in attachments.items():
                for k, v in attachment.items():
                    b64 = "".join(v) if isinstance(v, list) else v
                    try:
                        img_bytes = base64.b64decode(b64)
                        # infer ext from key
                        ext = k.split("/")[-1] if "/" in k else "png"
                        images.append((name, img_bytes, ext))
                    except Exception:
                        pass
        return "\n".join(text_parts), images

    # ---------- python file read ----------
    def read_py(self, file_bytes):
        try:
            txt = file_bytes.decode("utf-8", errors="ignore")
            return txt
        except Exception as e:
            return ""

    # ---------- OCR (Textract primary, pytesseract fallback) ----------
    def ocr_image_with_textract(self, image_bytes):
        """
        Use AWS Textract detect_document_text for OCR. Returns aggregated text.
        """
        if not self.TEXTRACT:
            return None
        try:
            response = self.TEXTRACT.detect_document_text(Document={'Bytes': image_bytes})
            # aggregate lines
            lines = [blk.get("Text", "") for blk in response.get("Blocks", []) if blk.get("BlockType") == "LINE"]
            return "\n".join(lines)
        except Exception as e:
            print(f"Textract OCR failed: {e}")
            return None

    def ocr_image_with_pytesseract(self, image_bytes):
        if not PYTESSACT_AVAILABLE:
            return None
        try:
            img = Image.open(io.BytesIO(image_bytes))
            txt = pytesseract.image_to_string(img)
            return txt
        except Exception as e:
            print(f"pytesseract OCR failed: {e}")
            return None

    def ocr_image(self, image_bytes):
        """
        Try Textract, fallback to pytesseract if installed. Returns text (possibly empty string).
        """
        txt = None
        if self.TEXTRACT:
            txt = self.ocr_image_with_textract(image_bytes)
        if not txt and PYTESSACT_AVAILABLE:
            txt = self.ocr_image_with_pytesseract(image_bytes)
        return txt or ""

    # ---------- GPT output cleaner ----------
    def cleaner(self, output={}):
        try:
            raw_text = output['evaluation'].strip()
            if raw_text.startswith("```json"):
                raw_text = raw_text[7:]
            if raw_text.endswith("```"):
                raw_text = raw_text[:-3]
            cleaned_json = json.loads(raw_text)
            return cleaned_json
        except Exception as e:
            print(f"JSON Parsing Failed: {e}")
            return {"error": "Failed to parse GPT output", "raw": output}

    # ---------- GPT evaluation ----------
    def evaluate_entire_assignment(self, content):
        prompt = f"""
        You are a strict assignment evaluator who excels in every known technology.

        The student has submitted an assignment containing multiple questions and answers . Please:
        1. Identify all distinct questions and their answers.
        2. For each, return feedback, a score from 0 to 10, and whether it's AI-generated or not.

        Respond **only in this JSON format**:
        {{
        "evaluation": [
            {{
            "Q1": {{
                "Question1": "...",
                "feedback": "...",
                "score": 0-10,
                "verdict": "AI-generated" or "Not AI-generated"
            }}
            }},
            ...
        ]
        }}

        Input:
        \"\"\"{content}\"\"\"
        """
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are a Python assignment evaluator."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                max_tokens=4096
            )
            result = response.model_dump()
            return result["choices"][0]["message"]["content"]
        except Exception as e:
            print(f"OpenAI Evaluation Failed: {e}")
            return json.dumps({"error": "Evaluation failed", "details": str(e)})

    # ---------- chunking & tokens ----------
    def chunk_text(self, text, max_words=3000):
        words = text.split()
        chunks = []
        for i in range(0, len(words), max_words):
            chunks.append(" ".join(words[i:i+max_words]))
        return chunks

    def count_tokens(self, text, model="gpt-4o"):
        encoding = tiktoken.encoding_for_model(model)
        return len(encoding.encode(text))

    # ---------- verification ----------
    def verify_evaluation(self, eval_text):
        MAX_TOKENS = 100000
        token_count = self.count_tokens(eval_text)

        if token_count > MAX_TOKENS:
            eval_text = eval_text[:50000]  # Naive trim

        prompt = f"""
        You are a verification expert.
        Does this JSON include question, feedback, score (0–10), and verdict (AI or Not AI)?
        Return True if valid, else False and why.

        {eval_text}
        """
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You're a verification expert."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=1200
        )
        result = response.model_dump()
        # result may have choices; safely extract
        try:
            verdict = result["choices"][0]["message"]["content"].strip()
        except Exception:
            verdict = ""
        return (True, "") if verdict.lower().startswith("true") else (False, verdict or result)

    # ---------- sorting ----------
    def sort_evaluation(self, evaluation_list):
        def extract_q_num(item):
            key = list(item.keys())[0]
            # remove non-numeric characters and convert
            try:
                return int(''.join(filter(str.isdigit, key)))
            except Exception:
                return 0
        return sorted(evaluation_list, key=extract_q_num)

    # ---------- PDF decorator & generation ----------
    def decorator(self, canvas_obj, doc):
        width, height = A4
        logo_path = "image.png"
        try:
            logo = ImageReader(logo_path)
            logo_width, logo_height = 200, 200
            x = (width - logo_width) / 2
            y = (height - logo_height) / 2

            canvas_obj.saveState()
            try:
                canvas_obj.setFillAlpha(0.1)
            except AttributeError:
                pass
            canvas_obj.drawImage(logo, x, y, width=logo_width, height=logo_height, mask='auto')
            canvas_obj.restoreState()
        except Exception as e:
            # If logo not found or fails, continue without it
            print(f"Logo draw skipped: {e}")

        canvas_obj.setStrokeColor(colors.grey)
        canvas_obj.setLineWidth(2)
        canvas_obj.rect(10, 10, width - 20, height - 20)

    def json_to_pdf_report(self, data, student_name, assignment_name, output_path="report.pdf"):
        pdfmetrics.registerFont(TTFont('DejaVuSans', 'dejavu-sans/DejaVuSans.ttf'))
        pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', 'dejavu-sans/DejaVuSans-Bold.ttf'))
        pdfmetrics.registerFontFamily('DejaVuSans',
                                     normal='DejaVuSans',
                                     bold='DejaVuSans-Bold')

        doc = SimpleDocTemplate(output_path, pagesize=A4,
                                rightMargin=1.75 * cm, leftMargin=1.75 * cm,
                                topMargin=2 * cm, bottomMargin=2 * cm)

        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name='Normal_DejaVu', fontName='DejaVuSans', fontSize=10, leading=12))
        styles.add(ParagraphStyle(name='mHeading', fontName='DejaVuSans-Bold', fontSize=14, alignment=TA_CENTER, spaceAfter=12))
        styles.add(ParagraphStyle(name='SubHeading', fontName='DejaVuSans-Bold', fontSize=11, spaceAfter=10, textColor=colors.HexColor("#2E86C1")))

        elements = [
            Paragraph("Assignment Report", styles['mHeading']),
            Paragraph(f"Student Mail: <font color='black'>{student_name}</font>", styles['SubHeading']),
            Paragraph(f"Assignment Name: <font color='black'>{assignment_name}</font>", styles['SubHeading']),
            Spacer(1, 12)
        ]

        page_width, _ = A4
        available_width = page_width - doc.leftMargin - doc.rightMargin
        col_widths = [available_width * 0.35, available_width * 0.40, available_width * 0.15, available_width * 0.10]

        table_data = [
            [Paragraph(f"<b>{col}</b>", styles['Normal_DejaVu']) for col in ["Question", "Feedback", "Verdict", "Score"]]
        ]

        data["evaluation"] = self.sort_evaluation(data["evaluation"])

        for item in data["evaluation"]:
            for _, evaluation in item.items():
                question_key = next((k for k in evaluation if k.lower().startswith("question")), None)
                row = [
                    Paragraph(evaluation.get(question_key, ""), styles['Normal_DejaVu']),
                    Paragraph(evaluation.get("feedback", ""), styles['Normal_DejaVu']),
                    Paragraph(evaluation.get("verdict", ""), styles['Normal_DejaVu']),
                    Paragraph(str(evaluation.get("score", "")), styles['Normal_DejaVu']),
                ]
                table_data.append(row)

        # append summary row if present in data
        if "summary" in data:
            summary = data["summary"]
            table_data.append([
                Paragraph("<b>Overall Summary</b>", styles['Normal_DejaVu']),
                Paragraph(summary.get("text", ""), styles['Normal_DejaVu']),
                Paragraph("", styles['Normal_DejaVu']),
                Paragraph(summary.get("average_score", ""), styles['Normal_DejaVu'])
            ])

        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.orange),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))

        elements.append(table)
        doc.build(elements, onFirstPage=self.decorator, onLaterPages=self.decorator)

    # ---------- upload PDF ----------
    def upload_pdf_to_s3(self, file_buffer, filename, metadata=None):
        content_type = "application/pdf"
        extra_args = {'ContentType': content_type}
        if metadata:
            extra_args['Metadata'] = {str(k): str(v) for k, v in metadata.items()}
        try:
            file_buffer.seek(0)
            self.S3_CLIENT.upload_fileobj(
                file_buffer,
                self.S3_BUCKET_NAME,
                filename,
                ExtraArgs=extra_args
            )
            return f"https://{self.S3_BUCKET_NAME}.s3.amazonaws.com/{filename}"
        except Exception as e:
            return {"error": "S3 upload failed", "details": str(e)}

    # ---------- DynamoDB student lookup ----------
    def get_student_data(self, email):
        if not email:
            return None
        try:
            resp = self.STUDENTS_TABLE.get_item(Key={"email": email})
            return resp.get("Item")
        except Exception as e:
            print(f"Failed to fetch student data for {email}: {e}")
            return None

    # ---------- DynamoDB class lookup ----------
    def get_class_data(self, class_code):
        if not class_code:
            return None
        try:
            resp = self.CLASS_TABLE.get_item(Key={"class_code": class_code})
            return resp.get("Item")
        except Exception as e:
            print(f"Failed to fetch class data for {class_code}: {e}")
            return None

    # ---------- Get assignment metadata (from generated-assignments bucket) ----------
    def get_assignment_metadata(self, assignment_name):
        try:
            resp = self.S3_CLIENT.head_object(Bucket=self.GENERATED_ASSIGNMENTS_BUCKET, Key=assignment_name)
            metadata = resp.get("Metadata", {})
            # metadata keys are lowercase in S3
            return {
                "class_code": metadata.get("class_code"),
                "deadline": metadata.get("deadline"),
                "extra_metadata": metadata
            }
        except ClientError as e:
            code = e.response.get("Error", {}).get("Code", "")
            print(f"Error fetching assignment metadata ({assignment_name}): {code} - {e}")
            return None
        except Exception as e:
            print(f"Unexpected error fetching assignment metadata: {e}")
            return None

    # ---------- parse ISO datetime to UTC-aware datetime for comparison ----------
    @staticmethod
    def parse_iso_datetime_to_utc(dt_str):
        """
        Accepts ISO datetime string. Returns naive UTC datetime for comparison with datetime.datetime.utcnow()
        If dt_str contains timezone, convert to UTC and return naive UTC.
        """
        if not dt_str:
            return None
        try:
            # Python 3.11+ fromisoformat supports offset; for older versions this still fairly robust
            dt = datetime.datetime.fromisoformat(dt_str)
        except Exception:
            # Fall back to common formats
            try:
                dt = datetime.datetime.strptime(dt_str, "%Y-%m-%dT%H:%M:%S")
            except Exception:
                try:
                    dt = datetime.datetime.strptime(dt_str, "%Y-%m-%d")
                    # set to end of day
                    dt = dt.replace(hour=23, minute=59, second=59)
                except Exception:
                    return None

        # If timezone-aware, convert to UTC and then drop tzinfo
        if dt.tzinfo is not None:
            dt_utc = dt.astimezone(datetime.timezone.utc).replace(tzinfo=None)
            return dt_utc
        # naive => assume UTC
        return dt

    def send_student_email(self,to_email, subject, body):
        smtp_server = os.getenv("SMTP_SERVER")
        smtp_port = int(os.getenv("SMTP_PORT") or 587)
        smtp_user = os.getenv("SMTP_USER")
        smtp_pass = os.getenv("SMTP_PASS")
        from_email = os.getenv("FROM_EMAIL")

        msg = MIMEMultipart()
        msg["From"] = from_email
        msg["To"] = to_email
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "plain"))

        try:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(smtp_user, smtp_pass)
                server.sendmail(from_email, to_email, msg.as_string())
            print(f"📧 Mail sent successfully to {to_email}")
        except Exception as e:
            print(f"❌ Failed to send email: {e}")

    # ---------- Excel update/creation (updated to include class table fields & summary) ----------
    def update_assignment_excel(self, assignment_name, student_data, evaluation, deadline, class_data=None, bucket=None):
        bucket = bucket or self.EXCEL_BUCKET

        # Build filename: college_classcode_assignment.xlsx
        college = (student_data.get("college_name") if student_data else "unknownCollege") or "unknownCollege"
        college_safe = college.replace(" ", "_")
        class_code = (student_data.get("class_code") if student_data else (class_data.get("class_code") if class_data else "unknownClass")) or "unknownClass"
        safe_assignment = assignment_name.replace(" ", "_")
        key = f"{college_safe}_{class_code}_{safe_assignment}.xlsx"

        # Parse deadline to compare
        if isinstance(deadline, str):
            parsed_deadline = self.parse_iso_datetime_to_utc(deadline)
        elif isinstance(deadline, datetime.datetime):
            parsed_deadline = deadline
        else:
            parsed_deadline = None

        if parsed_deadline is None:
            return {"error": "Invalid or missing deadline metadata. Excel not updated.", "key": key}

        now_utc = datetime.datetime.utcnow()
        if now_utc > parsed_deadline:
            return {"error": f"Deadline passed for {assignment_name}. Excel is locked.", "key": key}

        # Try to load existing workbook from S3
        wb = None
        ws = None
        try:
            obj = self.S3_CLIENT.get_object(Bucket=bucket, Key=key)
            file_bytes = io.BytesIO(obj["Body"].read())
            wb = load_workbook(file_bytes)
            ws = wb.active
        except ClientError as e:
            err_code = e.response.get("Error", {}).get("Code", "")
            # If not found, create new workbook
            if err_code in ("NoSuchKey", "404", "NotFound", "NoSuchBucket"):
                wb = Workbook()
                ws = wb.active
                ws.title = "Submissions"
                ws.append([
                    "Name", "Email", "Phone", "College", "College Email", "Class Code", "Department",
                    "Question", "Feedback", "Verdict", "Score"
                ])
            else:
                print(f"S3 get_object failed for {key}: {e}")
                return {"error": "Failed to load existing Excel", "details": str(e)}
        except Exception as e:
            # Some other error -> create workbook
            print(f"Unexpected error loading workbook from S3: {e}")
            wb = Workbook()
            ws = wb.active
            ws.title = "Submissions"
            ws.append([
                "Name", "Email", "Phone", "College", "College Email", "Class Code", "Department",
                "Question", "Feedback", "Verdict", "Score"
            ])

        # Append evaluation rows: one row per question
        try:
            eval_list = evaluation.get("evaluation", [])
            total_score = 0.0
            q_count = 0
            ai_count = 0
            for item in eval_list:
                # item is like {"Q1": { ... }}
                for _, qdata in item.items():
                    question_key = next((k for k in qdata if k.lower().startswith("question")), None)
                    score_val = qdata.get("score", "")
                    try:
                        score_num = float(score_val)
                    except Exception:
                        score_num = 0.0
                    total_score += score_num
                    q_count += 1
                    if qdata.get("verdict", "").lower().startswith("ai"):
                        ai_count += 1

                    ws.append([
                        student_data.get("name", "") if student_data else "",
                        student_data.get("email", "") if student_data else "",
                        student_data.get("phone", "") if student_data else "",
                        student_data.get("college_name", "") if student_data else "",
                        (class_data.get("college_email") if class_data else "") or "",
                        class_code,
                        (class_data.get("department") if class_data else student_data.get("department","") if student_data else "") or "",
                        qdata.get(question_key, "") if question_key else "",
                        qdata.get("feedback", ""),
                        qdata.get("verdict", ""),
                        qdata.get("score", "")
                    ])
            # Append summary row
            avg_score = round(total_score / q_count, 2) if q_count else ""
            summary_text = f"Total Questions: {q_count}, AI-detected: {ai_count}"
            ws.append([
                "","", "", "", "", "", "Overall Summary",
                summary_text, "", "", avg_score
            ])
        except Exception as e:
            print(f"Failed to append evaluation rows: {e}")
            return {"error": "Appending rows failed", "details": str(e)}

        # Save workbook to buffer and upload back to S3
        try:
            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            self.S3_CLIENT.upload_fileobj(
                buffer,
                bucket,
                key,
                ExtraArgs={"ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
            )
            file_url = f"https://{bucket}.s3.amazonaws.com/{key}"
            return {"message": "Excel updated successfully", "fileUrl": file_url, "key": key}
        except Exception as e:
            print(f"Failed to upload Excel to S3: {e}")
            return {"error": "Excel upload failed", "details": str(e)}

    # ---------- New: Evaluate a generic file by type ----------
    def prepare_content_from_file(self, file_bytes, filename):
        """
        Detect file type by extension and assemble a textual content string summarizing:
        - extracted text (pdf/docx/ipynb/py)
        - embedded images OCR results and small notes about them
        Returns combined text content string and list of extracted images (for possible later use)
        """
        lname = filename.lower()
        images_collected = []
        content_parts = []

        try:
            if lname.endswith(".pdf"):
                # text
                txt = self.read_pdf(file_bytes)
                content_parts.append("=== Extracted PDF Text ===\n" + (txt or ""))
                # images
                try:
                    imgs = self.extract_images_from_pdf(file_bytes)
                    for img_name, img_bytes, ext in imgs:
                        ocr = self.ocr_image(img_bytes)
                        content_parts.append(f"--- Image: {img_name} (from PDF) ---\n{(ocr or '[no OCR text found]')}\n")
                        images_collected.append((img_name, img_bytes, ext))
                except Exception as e:
                    print(f"PDF image extraction failed: {e}")

            elif lname.endswith(".ipynb") or lname.endswith(".ipynb.json"):
                text, imgs = self.read_ipynb(file_bytes)
                content_parts.append("=== Notebook Content ===\n" + (text or ""))
                for img_name, img_bytes, ext in imgs:
                    ocr = self.ocr_image(img_bytes)
                    content_parts.append(f"--- Notebook Image: {img_name} ---\n{(ocr or '[no OCR text found]')}\n")
                    images_collected.append((img_name, img_bytes, ext))

            elif lname.endswith(".docx") or lname.endswith(".doc"):
                text, imgs = self.read_docx(file_bytes)
                content_parts.append("=== Document Text ===\n" + (text or ""))
                for img_name, img_bytes, ext in imgs:
                    ocr = self.ocr_image(img_bytes)
                    content_parts.append(f"--- Document Image: {img_name} ---\n{(ocr or '[no OCR text found]')}\n")
                    images_collected.append((img_name, img_bytes, ext))

            elif lname.endswith(".py"):
                txt = self.read_py(file_bytes)
                content_parts.append("=== Python File ===\n" + (txt or ""))

            # single image files
            elif any(lname.endswith(ext) for ext in [".png", ".jpg", ".jpeg", ".bmp", ".tiff"]):
                # treat as image
                ocr = self.ocr_image(file_bytes)
                content_parts.append("=== Image Submission ===\n" + (ocr or "[no OCR text found]"))
                images_collected.append((filename, file_bytes, lname.split('.')[-1]))

            else:
                # fallback: try to read as text and include a note
                try:
                    txt = file_bytes.decode("utf-8", errors="ignore")
                    content_parts.append("=== Fallback Text Read ===\n" + (txt or ""))
                except Exception:
                    content_parts.append("=== Unsupported file type. Content included as binary summary only. ===")

        except Exception as e:
            print(f"prepare_content_from_file error: {e}")
            content_parts.append(f"Note: failed to fully parse file due to: {e}")

        combined = "\n\n".join(content_parts)
        return combined, images_collected


# ---------- Evaluator (Lambda-like handler) ----------
class Evaluator:
    def evl(self, event):
        obj = Processor()
        body = event['request']['body']
        if isinstance(body, str):
            body = json.loads(body)

        # presigned or direct file URL posted in body
        presigned_url = body.get("file_path")
        if not presigned_url:
            return {"statusCode": 400, "body": json.dumps({"error": "No file_path provided in request body"})}

        print("📥 Downloading file from S3...")
        try:
            resp = requests.get(presigned_url)
            resp.raise_for_status()
        except Exception as e:
            print(f"Failed to download file from presigned URL: {e}")
            return {"statusCode": 400, "body": json.dumps({"error": "Failed to download file", "details": str(e)})}

        file_bytes = io.BytesIO(resp.content).getvalue()
        filename = body.get("assignment_name") or os.path.basename(presigned_url)

        print("🔎 Preparing content (text + image OCR) from submitted file...")
        try:
            combined_text, images = obj.prepare_content_from_file(file_bytes, filename)
            if not combined_text.strip():
                combined_text = "[No textual content could be extracted from the submitted file. Images and diagrams (if any) included below or OCR attempted.]"
        except Exception as e:
            print(f"File parsing failed: {e}")
            return {"statusCode": 500, "body": json.dumps({"error": "Failed to parse submitted file", "details": str(e)})}

        print("🔍 Chunking text for evaluation...")
        chunks = obj.chunk_text(combined_text)
        evaluations = []

        print("🧠 Running GPT evaluations...")
        for idx, chunk in enumerate(chunks):
            print(f"→ Evaluating chunk {idx+1}/{len(chunks)}...")
            chunk_output = obj.evaluate_entire_assignment(chunk)
            cleaned = obj.cleaner({"evaluation": chunk_output})
            # if cleaner returns error, stop and return
            if "error" in cleaned:
                print(f"Cleaner failed for chunk {idx+1}: {cleaned}")
                return {"statusCode": 500, "body": json.dumps({"error": "Cleaner failed", "details": cleaned})}
            evaluations.extend(cleaned.get("evaluation", []))

        cleaned_data = {"evaluation": evaluations}

        print("✅ Verifying output...")
        try:
            verified, reason = obj.verify_evaluation(json.dumps(cleaned_data))
        except Exception as e:
            print(f"Verification call failed: {e}")
            return {"statusCode": 500, "body": json.dumps({"error": "Verification failed", "details": str(e)})}

        if not verified:
            print(f"❌ JSON verification failed: {reason}")
            return {"statusCode": 400, "body": json.dumps({"error": reason})}

        # Add an overall summary: average score, total questions, AI-count -- saved into cleaned_data["summary"]
        try:
            eval_list = cleaned_data.get("evaluation", [])
            total = 0.0
            qcount = 0
            ai_count = 0
            for item in eval_list:
                for _, qdata in item.items():
                    sc = qdata.get("score", 0)
                    try:
                        total += float(sc)
                    except Exception:
                        total += 0.0
                    qcount += 1
                    if qdata.get("verdict", "").lower().startswith("ai"):
                        ai_count += 1
            avg = round(total / qcount, 2) if qcount else ""
            cleaned_data["summary"] = {
                "total_questions": qcount,
                "average_score": avg,
                "ai_detected": ai_count,
                "text": f"Total Questions: {qcount}. Average score: {avg}. AI-detected answers: {ai_count}."
            }
        except Exception as e:
            print(f"Failed to compute summary: {e}")

        print("📝 Generating PDF...")
        output_pdf_path = "evaluation_report.pdf"
        try:
            obj.json_to_pdf_report(cleaned_data, body.get("student_name", ""), body.get("assignment_name", ""), output_pdf_path)
        except Exception as e:
            print(f"PDF generation failed: {e}")
            return {"statusCode": 500, "body": json.dumps({"error": "PDF generation failed", "details": str(e)})}

        print("☁️ Uploading PDF to S3...")
        try:
            with open(output_pdf_path, "rb") as f:
                buffer = io.BytesIO(f.read())
                std = body.get("student_name", "").split('@')[0]
                assn = body.get("assignment_name", "").split('.')[0]
                pdf_filename = f"{std}_{assn}.pdf"
                pdf_url = obj.upload_pdf_to_s3(
                    buffer,
                    filename=pdf_filename,
                    metadata={
                        "email": body.get("student_name"),
                        "assignment_name": body.get("assignment_name"),
                        "class_code": body.get("class_code", "")
                    }
                )
        except Exception as e:
            print(f"S3 PDF upload failed: {e}")
            return {"statusCode": 500, "body": json.dumps({"error": "PDF upload failed", "details": str(e)})}

        s3_client = boto3.client("s3", region_name=os.getenv("AWS_REGION", "us-east-1"))
        presigned_url = s3_client.generate_presigned_url(
            "get_object",
            Params={
                "Bucket": obj.S3_BUCKET_NAME,
                "Key": f"{std}_{assn}.pdf"
            },
            ExpiresIn=604800  # 7 days
        )

        # Send mail to student
        try:
            obj.send_student_email(
                to_email=body.get("student_name"),  # student's email
                subject=f"Your Evaluation Report for {body.get('assignment_name')}",
                body=f"""Hello,

Your assignment has been evaluated.

You can download your report from the link below (valid for 7 days):
{presigned_url}

Best regards,
Assignment Evaluator
"""
            )
        except Exception as e:
            print(f"Failed to send email: {e}")

        # ---------- NEW: Excel update ----------
        print("📥 Fetching assignment metadata (deadline & class_code)...")
        assignment_name = body.get("assignment_name")
        if not assignment_name:
            return {"statusCode": 400, "body": json.dumps({"error": "assignment_name missing in request body"})}

        assignment_meta = obj.get_assignment_metadata(assignment_name)
        if not assignment_meta:
            # if metadata fetch failed, still return success for evaluation but note metadata fetch failure
            return {
                "statusCode": 200,
                "headers": {
                    "Access-Control-Allow-Origin": "*",
                    "Access-Control-Allow-Methods": "OPTIONS,POST,GET",
                    "Access-Control-Allow-Headers": "Content-Type"
                },
                "body": json.dumps({
                    "message": "Evaluation completed and PDF uploaded, but failed to fetch assignment metadata for Excel update.",
                    "fileUrl": pdf_url
                })
            }

        deadline_str = assignment_meta.get("deadline")
        # parse deadline
        parsed_deadline = obj.parse_iso_datetime_to_utc(deadline_str)
        if parsed_deadline is None:
            # same fallback: return success for evaluation but warn about deadline parse issue
            return {
                "statusCode": 200,
                "headers": {
                    "Access-Control-Allow-Origin": "*",
                    "Access-Control-Allow-Methods": "OPTIONS,POST,GET",
                    "Access-Control-Allow-Headers": "Content-Type"
                },
                "body": json.dumps({
                    "message": "Evaluation completed and PDF uploaded, but assignment deadline metadata was invalid. Excel not updated.",
                    "fileUrl": pdf_url
                })
            }

        # Fetch student data from DynamoDB using the email (student_name field contains email)
        student_email = body.get("student_name")  # you use student_name as email in previous code
        student_data = obj.get_student_data(student_email)
        if not student_data:
            # If no student data, still return success of PDF but report missing student entry
            return {
                "statusCode": 200,
                "headers": {
                    "Access-Control-Allow-Origin": "*",
                    "Access-Control-Allow-Methods": "OPTIONS,POST,GET",
                    "Access-Control-Allow-Headers": "Content-Type"
                },
                "body": json.dumps({
                    "message": "Evaluation completed and PDF uploaded, but student record not found in DynamoDB. Excel not updated.",
                    "fileUrl": pdf_url
                })
            }

        # Fetch class data (for department and college mail)
        class_code = assignment_meta.get("class_code") or student_data.get("class_code")
        class_data = obj.get_class_data(class_code) if class_code else None

        # Update excel
        print("📊 Updating Excel tracker in S3...")
        try:
            excel_result = obj.update_assignment_excel(
                assignment_name=assignment_name,
                student_data=student_data,
                evaluation=cleaned_data,
                deadline=parsed_deadline,
                class_data=class_data,
                bucket=obj.EXCEL_BUCKET
            )
        except Exception as e:
            print(f"Excel update failed unexpectedly: {e}")
            excel_result = {"error": "Excel update threw exception", "details": str(e)}

        # Build final response: include pdf_url and excel_result (if any)
        response_body = {
            "message": "Evaluation report generated and uploaded successfully.",
            "fileUrl": pdf_url,
            "excel_result": excel_result
        }

        return {
            "statusCode": 200,
            "headers": {
                "Access-Control-Allow-Origin": "*",
                "Access-Control-Allow-Methods": "OPTIONS,POST,GET",
                "Access-Control-Allow-Headers": "Content-Type"
            },
            "body": json.dumps(response_body)
        }
