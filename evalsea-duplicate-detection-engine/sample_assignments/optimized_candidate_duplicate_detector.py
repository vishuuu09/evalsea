import os
import hashlib
import json
import ast
import difflib

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

import PyPDF2
import docx
import nbformat

# =========================================================
# CONFIG
# =========================================================
BASE_DIR = "/Users/mohammadafsheen/Desktop/sample_assignments"  # ✅ CHANGE IF NEEDED
SIMILARITY_THRESHOLD = 0.75   # Only report above this
FAST_SIGNATURE_WORDS = 180    # Controls bucket tightness

# =========================================================
# FILE EXTRACTION
# =========================================================
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
        return "\n".join([p.text for p in d.paragraphs])

    if path.endswith(".ipynb"):
        nb = nbformat.read(path, as_version=4)
        return "\n".join(cell.source for cell in nb.cells)

    if path.endswith(".py") or path.endswith(".txt"):
        return open(path, "r", errors="ignore").read()

    return ""

# =========================================================
# FAST SIGNATURE (CANDIDATE FILTER)
# =========================================================
def fast_signature(text):
    tokens = text.lower().split()[:FAST_SIGNATURE_WORDS]
    base = " ".join(tokens)
    return hashlib.md5(base.encode()).hexdigest()

# =========================================================
# AST STRUCTURE SIMILARITY (FOR PYTHON)
# =========================================================
def ast_similarity(code1, code2):
    try:
        t1 = ast.dump(ast.parse(code1))
        t2 = ast.dump(ast.parse(code2))
        return difflib.SequenceMatcher(None, t1, t2).ratio()
    except:
        return 0.0

# =========================================================
# LOAD LOCAL SUBMISSIONS
# =========================================================
def load_submissions():
    print("🔍 Loading local submissions...")

    submissions = {}

    for student in os.listdir(BASE_DIR):
        student_path = os.path.join(BASE_DIR, student)

        if not os.path.isdir(student_path):
            continue

        for file in os.listdir(student_path):
            full_path = os.path.join(student_path, file)

            submissions[student] = {
                "file": file,
                "path": full_path,
                "text": extract_text(full_path),
                "is_code": file.endswith(".py")
            }
            break

    print("✅ Total submissions found:", len(submissions))
    return submissions

# =========================================================
# BUCKETING (LIKELY MATCH GROUPS)
# =========================================================
def build_candidate_buckets(submissions):
    print("⚡ Building fast candidate buckets...")
    buckets = {}

    for sid, data in submissions.items():
        sig = fast_signature(data["text"])
        buckets.setdefault(sig, []).append(sid)

    print("✅ Total buckets formed:", len(buckets))
    return buckets

# =========================================================
# SEMANTIC SIMILARITY (TF-IDF)
# =========================================================
def compute_semantic_matrix(texts):
    vectorizer = TfidfVectorizer(stop_words="english")
    tfidf = vectorizer.fit_transform(texts)
    return cosine_similarity(tfidf)

# =========================================================
# MAIN OPTIMIZED DETECTOR
# =========================================================
def run_optimized_detector():
    submissions = load_submissions()
    buckets = build_candidate_buckets(submissions)

    final_results = []

    print("🤖 Running deep similarity ONLY inside candidate buckets...\n")

    for bucket_key, student_list in buckets.items():

        # ✅ Skip single-file buckets instantly
        if len(student_list) < 2:
            continue

        texts = [submissions[s]["text"] for s in student_list]
        sim_matrix = compute_semantic_matrix(texts)

        # ✅ Compare ONLY inside this small bucket
        for i in range(len(student_list)):
            for j in range(i + 1, len(student_list)):
                score = sim_matrix[i][j]

                if score >= SIMILARITY_THRESHOLD:
                    s1 = student_list[i]
                    s2 = student_list[j]

                    ast_score = 0.0
                    if submissions[s1]["is_code"] and submissions[s2]["is_code"]:
                        ast_score = ast_similarity(
                            submissions[s1]["text"],
                            submissions[s2]["text"]
                        )

                    final_results.append({
                        "student_1": s1,
                        "student_2": s2,
                        "file_1": submissions[s1]["file"],
                        "file_2": submissions[s2]["file"],
                        "semantic_similarity": round(float(score), 4),
                        "ast_similarity": round(float(ast_score), 4),
                        "verdict": "DUPLICATE"
                    })

    print("\n✅ FINAL OPTIMIZED DUPLICATE REPORT\n")
    print(json.dumps(final_results, indent=2))

# =========================================================
# ENTRY POINT
# =========================================================
if __name__ == "__main__":
    run_optimized_detector()
