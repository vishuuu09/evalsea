import json
# import boto3
import tempfile
import os
import hashlib
import numpy as np
from nltk.tokenize import word_tokenize
from simhash import Simhash
from datasketch import MinHash
import PyPDF2
import docx
import nbformat
import ast
import difflib
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# -----------------------------
# CONFIG
# -----------------------------
BASE_DIR = "/Users/mohammadafsheen/Desktop/lms_parent_portal /evalsea-duplicate-detection-engine/sample_assignments"


WEIGHTS = {
    "semantic": 0.40,
    "minhash": 0.25,
    "simhash": 0.20,
    "sha256": 0.15,
    "ast": 0.10   # only applied if both are Python
}

# -----------------------------
# FILE EXTRACTION
# -----------------------------
def extract_text(path):
    if path.endswith(".pdf"):
        text = ""
        reader = PyPDF2.PdfReader(open(path, "rb"))
        for p in reader.pages:
            t = p.extract_text()
            if t:
                text += t + "\n"
        return text.strip()

    if path.endswith(".docx"):
        d = docx.Document(path)
        return "\n".join([p.text for p in d.paragraphs]).strip()

    if path.endswith(".ipynb"):
        nb = nbformat.read(path, as_version=4)
        content = []
        for cell in nb.cells:
            content.append(cell.source)
        return "\n".join(content)

    if path.endswith(".py") or path.endswith(".txt"):
        return open(path, "r", errors="ignore").read()

    return ""

# -----------------------------
# PREPROCESSING
# -----------------------------
def preprocess(text):
    text = text.lower()
    return " ".join(word_tokenize(text))

# -----------------------------
# SHA-256
# -----------------------------
def sha256_hash(text):
    return hashlib.sha256(text.encode()).hexdigest()

# -----------------------------
# SIMHASH
# -----------------------------
def simhash_value(text):
    return Simhash(text).value

def simhash_similarity(v1, v2):
    return 1 - (bin(v1 ^ v2).count("1") / 64)

# -----------------------------
# MINHASH
# -----------------------------
def minhash_sig(text):
    m = MinHash(num_perm=128)
    for t in text.split():
        m.update(t.encode("utf8"))
    return m

# -----------------------------
# AST LOGIC MATCH (PYTHON)
# -----------------------------
def ast_similarity(c1, c2):
    try:
        t1 = ast.dump(ast.parse(c1))
        t2 = ast.dump(ast.parse(c2))
        return difflib.SequenceMatcher(None, t1, t2).ratio()
    except:
        return 0

# -----------------------------
# SEMANTIC EMBEDDING
# -----------------------------
def semantic_similarity(t1, t2):
    try:
        vectorizer = TfidfVectorizer()
        tfidf = vectorizer.fit_transform([t1, t2])
        return cosine_similarity(tfidf[0:1], tfidf[1:2])[0][0]
    except:
        return 0.0

# -----------------------------
# ENSEMBLE SCORING
# -----------------------------
def ensemble_similarity(t1, t2, is_code=False):
    p1, p2 = preprocess(t1), preprocess(t2)

    sha = 1.0 if sha256_hash(p1) == sha256_hash(p2) else 0
    simhash = simhash_similarity(simhash_value(p1), simhash_value(p2))
    minhash = minhash_sig(p1).jaccard(minhash_sig(p2))
    semantic = semantic_similarity(p1, p2)

    final = (
        WEIGHTS["semantic"] * semantic +
        WEIGHTS["minhash"] * minhash +
        WEIGHTS["simhash"] * simhash +
        WEIGHTS["sha256"] * sha
    )

    if is_code:
        final += WEIGHTS["ast"] * ast_similarity(t1, t2)

    return {
        "final": round(final, 4),
        "sha256": round(sha, 4),
        "simhash": round(simhash, 4),
        "minhash": round(minhash, 4),
        "semantic": round(semantic, 4),
    }

# -----------------------------
# AUTO DISCOVER ASSIGNMENT FOLDER
# -----------------------------
# def find_assignment_folder():
#     for root, dirs, files in os.walk(BASE_DIR):
#         for d in dirs:
#             if d.isdigit():
#                 return os.path.join(root, d)
#     return None

def find_assignment_folder():
    return BASE_DIR


# -----------------------------
# LOAD ALL STUDENT FILES
# -----------------------------
def load_submissions(assign_path):
    student_data = {}

    for student in os.listdir(assign_path):
        student_path = os.path.join(assign_path, student)
        if not os.path.isdir(student_path):
            continue

        for file in os.listdir(student_path):
            full = os.path.join(student_path, file)
            student_data[student] = {
                "file": file,
                "path": full,
                "text": extract_text(full),
                "is_code": file.endswith(".py")
            }
            break
    return student_data

# -----------------------------
# MAIN EXECUTION
# -----------------------------
def lambda_handler(event, context):
    """
    Expected Event:
    {
        "bucket": "submitted-assignments",
        "assignment_prefix": "testcollege/cse/2nd-year/a/class101/faculty-at-college-edu/202511080637/"
    }
    """

    try:
        bucket = event["bucket"]
        prefix = event["assignment_prefix"]

        print("Bucket:", bucket)
        print("Assignment Prefix:", prefix)

        # -----------------------------
        # STEP 1: LIST ALL FILES
        # -----------------------------
        paginator = s3.get_paginator("list_objects_v2")
        pages = paginator.paginate(Bucket=bucket, Prefix=prefix)

        student_files = {}

        for page in pages:
            for obj in page.get("Contents", []):
                key = obj["Key"]

                # Ignore folders
                if key.endswith("/"):
                    continue

                # Only allow these formats
                if not key.lower().endswith((".pdf", ".docx", ".py", ".ipynb", ".txt")):
                    continue

                parts = key.split("/")
                student_email = parts[-2]
                filename = parts[-1]

                local_path = os.path.join(tempfile.gettempdir(), filename)

                # -----------------------------
                # STEP 2: DOWNLOAD FILE
                # -----------------------------
                s3.download_file(bucket, key, local_path)

                # -----------------------------
                # STEP 3: EXTRACT TEXT
                # -----------------------------
                text = extract_text(local_path)

                student_files[student_email] = {
                    "file": filename,
                    "text": text,
                    "is_code": filename.endswith(".py")
                }

        if len(student_files) < 2:
            return {
                "statusCode": 200,
                "body": json.dumps({
                    "message": "Not enough submissions to check duplicates"
                })
            }

        # -----------------------------
        # STEP 4: DUPLICATE DETECTION
        # -----------------------------
        results = []
        students = list(student_files.keys())

        for i in range(len(students)):
            for j in range(i + 1, len(students)):
                s1, s2 = students[i], students[j]
                d1, d2 = student_files[s1], student_files[s2]

                score = ensemble_similarity(
                    d1["text"],
                    d2["text"],
                    is_code=d1["is_code"] and d2["is_code"]
                )

                verdict = (
                    "EXACT" if score["final"] >= 0.95 else
                    "HIGH" if score["final"] >= 0.80 else
                    "MEDIUM" if score["final"] >= 0.60 else
                    "LOW"
                )

                results.append({
                    "student_1": s1,
                    "student_2": s2,
                    "file_1": d1["file"],
                    "file_2": d2["file"],
                    "final_similarity": score["final"],
                    "sha256": score["sha256"],
                    "simhash": score["simhash"],
                    "minhash": score["minhash"],
                    "semantic": score["semantic"],
                    "verdict": verdict
                })

        # -----------------------------
        # STEP 5: RETURN RESPONSE
        # -----------------------------
        return {
            "statusCode": 200,
            "body": json.dumps({
                "bucket": bucket,
                "assignment_prefix": prefix,
                "total_students": len(student_files),
                "total_comparisons": len(results),
                "duplicate_results": results
            }, indent=2)
        }

    except Exception as e:
        return {
            "statusCode": 500,
            "body": json.dumps({
                "error": str(e)
            })
        }