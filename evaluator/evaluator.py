# 📌 Imports
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
load_dotenv()

client = OpenAI(api_key=os.getenv("OPEN_AI_API_KEY"))


class Processor:
    def read_pdf(self, file_bytes):
        text = ""
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for page in doc:
            text += page.get_text() + "\n"
        return text

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
                "score": 0–10,
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

    def chunk_text(self, text, max_words=3000):
        words = text.split()
        chunks = []
        for i in range(0, len(words), max_words):
            chunks.append(" ".join(words[i:i+max_words]))
        return chunks

    def count_tokens(self, text, model="gpt-4o"):
        encoding = tiktoken.encoding_for_model(model)
        return len(encoding.encode(text))

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
        verdict = response.choices[0].message.content.strip()
        return (True, "") if verdict.lower().startswith("true") else (False, result)

    def sort_evaluation(self, evaluation_list):
        def extract_q_num(item):
            key = list(item.keys())[0]
            return int(key[1:])
        return sorted(evaluation_list, key=extract_q_num)

    def decorator(self, canvas_obj, doc):
        width, height = A4
        logo_path = "image.png"
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
                                rightMargin=1.75*cm, leftMargin=1.75*cm,
                                topMargin=2*cm, bottomMargin=2*cm)

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

        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.orange),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))

        elements.append(table)
        doc.build(elements, onFirstPage=self.decorator, onLaterPages=self.decorator)

    

    S3_CLIENT = boto3.client(
        "s3",
        aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
            aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
        region_name="us-east-1"
    )
    S3_BUCKET_NAME = "evaluated-reports"

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

class Evaluator:
    def evl(self, event):
        obj = Processor()
        body = event['request']['body']
        if isinstance(body, str):
            body = json.loads(body)

        presigned_url = body.get("file_path")
        print("📥 Downloading file from S3...")
        resp = requests.get(presigned_url)
        resp.raise_for_status()

        file_bytes = io.BytesIO(resp.content)
        print("📄 Extracting text from PDF...")
        extracted_text = obj.read_pdf(file_bytes)

        print("🔍 Chunking text for evaluation...")
        chunks = obj.chunk_text(extracted_text)
        evaluations = []

        print("🧠 Running GPT evaluations...")
        for idx, chunk in enumerate(chunks):
            print(f"→ Evaluating chunk {idx+1}/{len(chunks)}...")
            chunk_output = obj.evaluate_entire_assignment(chunk)
            cleaned = obj.cleaner({"evaluation": chunk_output})
            evaluations.extend(cleaned.get("evaluation", []))

        cleaned_data = {"evaluation": evaluations}

        print("✅ Verifying output...")
        verified, reason = obj.verify_evaluation(json.dumps(cleaned_data))
        if not verified:
            print(f"❌ JSON verification failed: {reason}")
            return {"statusCode": 400, "body": json.dumps({"error": reason})}

        print("📝 Generating PDF...")
        output_pdf_path = "evaluation_report.pdf"
        obj.json_to_pdf_report(cleaned_data, body["student_name"], body["assignment_name"], output_pdf_path)

        print("☁️ Uploading to S3...")
        with open(output_pdf_path, "rb") as f:
            buffer = io.BytesIO(f.read())
            std = body.get("student_name").split('@')[0]
            assn = body.get("assignment_name").split('.')[0]
            url = obj.upload_pdf_to_s3(
                buffer,
                filename=f"{std}_{assn}.pdf",
                metadata={
                    "email": body.get("student_name"),
                    "assignment_name": body.get("assignment_name"),
                    "batch_name": body.get("batch_name")
                }
            )

        return {
            "statusCode": 200,
            "headers": {
                    "Access-Control-Allow-Origin": "*",
                    "Access-Control-Allow-Methods": "OPTIONS,POST,GET",
                    "Access-Control-Allow-Headers": "Content-Type"
                },
            "body": json.dumps({
                "message": "Evaluation report generated and uploaded successfully.",
                "fileUrl": url,
            })
        }
