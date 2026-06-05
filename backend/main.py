import os
import json
from datetime import datetime
from typing import List, Optional

import firebase_admin
from firebase_admin import credentials, firestore
from fastapi import FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from openpyxl import load_workbook

ADMIN_TOKEN = os.getenv("EDUCBT_ADMIN_TOKEN", "change-this-token")
ALLOWED_ORIGINS = os.getenv("EDUCBT_ALLOWED_ORIGINS", "*").split(",")

app = FastAPI(title="EduCBT Backend", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def init_firebase():
    if firebase_admin._apps:
        return firestore.client()

    service_json = os.getenv("FIREBASE_SERVICE_ACCOUNT_JSON")
    service_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")

    if service_json:
        cred = credentials.Certificate(json.loads(service_json))
    elif service_path:
        cred = credentials.Certificate(service_path)
    else:
        raise RuntimeError("Firebase credentials missing. Set FIREBASE_SERVICE_ACCOUNT_JSON or GOOGLE_APPLICATION_CREDENTIALS.")

    firebase_admin.initialize_app(cred)
    return firestore.client()


def verify_admin(token: Optional[str]):
    if not token or token != ADMIN_TOKEN:
        raise HTTPException(status_code=401, detail="Unauthorized admin request")


def normalize_header(value):
    return str(value or "").strip().lower().replace("_", " ")


def get_value(row, headers, names):
    wanted = [normalize_header(x) for x in names]
    for index, header in enumerate(headers):
        if normalize_header(header) in wanted:
            return row[index]
    return ""


VALID_SUBJECTS = {
    "Economics", "Government", "Commerce", "Accounting", "English Language", "Mathematics",
    "Biology", "Chemistry", "Physics", "Agricultural Science", "Civic Education", "CRS",
    "Basic Science", "Basic Technology", "Social Studies", "Business Studies", "Computer Studies",
    "Home Economics", "PHE"
}

VALID_CLASSES = {"JSS", "JSS1", "JSS2", "JSS3", "SS", "SS1", "SS2", "SS3", ""}
VALID_ANSWERS = {"A", "B", "C", "D"}


class Settings(BaseModel):
    schoolName: str = "EduCBT Portal"
    logoUrl: str = ""
    motto: str = "Knowledge and Excellence"
    schoolCode: str = "EDU"
    term: str = "First Term"
    session: str = "2025/2026"
    examType: str = "Quiz"
    questionsPerExam: int = 20
    timeMinutes: int = 15
    allowRetake: str = "yes"
    showCorrection: str = "yes"


@app.get("/health")
def health():
    return {"ok": True, "service": "EduCBT Backend", "time": datetime.utcnow().isoformat()}


@app.post("/admin/settings")
def save_settings(settings: Settings, x_admin_token: Optional[str] = Header(default=None)):
    verify_admin(x_admin_token)
    db = init_firebase()
    db.collection("settings").document("exam").set(settings.dict(), merge=True)
    return {"ok": True, "message": "Settings saved"}


@app.get("/admin/settings")
def get_settings(x_admin_token: Optional[str] = Header(default=None)):
    verify_admin(x_admin_token)
    db = init_firebase()
    doc = db.collection("settings").document("exam").get()
    return doc.to_dict() if doc.exists else Settings().dict()


@app.post("/admin/questions/import")
async def import_questions(
    target_class: str = Form(""),
    mode: str = Form("ALL"),
    file: UploadFile = File(...),
    x_admin_token: Optional[str] = Header(default=None),
):
    verify_admin(x_admin_token)

    if not file.filename.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        raise HTTPException(status_code=400, detail="Upload an Excel .xlsx file")

    contents = await file.read()
    temp_path = f"/tmp/{file.filename}"
    with open(temp_path, "wb") as f:
        f.write(contents)

    wb = load_workbook(temp_path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise HTTPException(status_code=400, detail="Empty Excel file")

    headers = [str(x or "").strip() for x in rows[0]]
    parsed = []

    for row_number, row in enumerate(rows[1:], start=2):
        subject = str(get_value(row, headers, ["Subject"] ) or "").strip()
        if mode != "ALL" and not subject:
            subject = mode

        class_level = str(get_value(row, headers, ["Class", "Class Level", "Level"] ) or target_class or "").strip().upper()
        question = str(get_value(row, headers, ["Question", "q"] ) or "").strip()
        options = [
            str(get_value(row, headers, ["A", "Option A"] ) or "").strip(),
            str(get_value(row, headers, ["B", "Option B"] ) or "").strip(),
            str(get_value(row, headers, ["C", "Option C"] ) or "").strip(),
            str(get_value(row, headers, ["D", "Option D"] ) or "").strip(),
        ]
        answer = str(get_value(row, headers, ["Correct Answer", "Answer"] ) or "").strip().upper()

        if not subject and not question:
            continue

        if subject not in VALID_SUBJECTS:
            raise HTTPException(status_code=400, detail=f"Invalid subject on row {row_number}: {subject}")
        if class_level not in VALID_CLASSES:
            raise HTTPException(status_code=400, detail=f"Invalid class on row {row_number}: {class_level}")
        if not question or any(not option for option in options):
            raise HTTPException(status_code=400, detail=f"Missing question/options on row {row_number}")
        if answer not in VALID_ANSWERS:
            raise HTTPException(status_code=400, detail=f"Invalid answer on row {row_number}: {answer}")

        parsed.append({
            "classLevel": class_level,
            "subject": subject,
            "q": question,
            "options": options,
            "answer": answer,
            "active": True,
            "createdAt": firestore.SERVER_TIMESTAMP,
        })

    if not parsed:
        raise HTTPException(status_code=400, detail="No valid questions found")

    db = init_firebase()
    pairs = {(item["classLevel"], item["subject"]) for item in parsed}

    old_docs = db.collection("quizQuestions").stream()
    batch = db.batch()
    ops = 0

    for doc in old_docs:
        data = doc.to_dict()
        pair = (str(data.get("classLevel", "")).upper(), data.get("subject", ""))
        if pair in pairs:
            batch.update(doc.reference, {"active": False})
            ops += 1
            if ops >= 400:
                batch.commit()
                batch = db.batch()
                ops = 0

    for item in parsed:
        ref = db.collection("quizQuestions").document()
        batch.set(ref, item)
        ops += 1
        if ops >= 400:
            batch.commit()
            batch = db.batch()
            ops = 0

    if ops:
        batch.commit()

    return {
        "ok": True,
        "imported": len(parsed),
        "groups": len(pairs),
        "message": f"Imported {len(parsed)} questions safely",
    }


@app.get("/admin/results")
def list_results(
    className: Optional[str] = None,
    subject: Optional[str] = None,
    term: Optional[str] = None,
    session: Optional[str] = None,
    x_admin_token: Optional[str] = Header(default=None),
):
    verify_admin(x_admin_token)
    db = init_firebase()
    query = db.collection("quizResults")
    if className:
        query = query.where("className", "==", className)
    if subject:
        query = query.where("subject", "==", subject)
    if term:
        query = query.where("term", "==", term)
    if session:
        query = query.where("session", "==", session)

    docs = query.limit(500).stream()
    results = []
    for doc in docs:
        item = doc.to_dict()
        item["id"] = doc.id
        results.append(item)
    return {"ok": True, "count": len(results), "results": results}
