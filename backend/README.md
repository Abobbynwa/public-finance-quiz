# EduCBT FastAPI Backend

This backend moves sensitive admin operations away from the browser.

## Features

- Protected admin settings endpoint
- Protected Excel question import endpoint
- Protected results listing endpoint
- Firebase Admin SDK server-side database writes

## Required Environment Variables

```bash
EDUCBT_ADMIN_TOKEN=change-to-a-long-secret-token
FIREBASE_SERVICE_ACCOUNT_JSON='{"type":"service_account",...}'
EDUCBT_ALLOWED_ORIGINS=https://public-finance-quiz.vercel.app
```

Alternative Firebase credential option:

```bash
GOOGLE_APPLICATION_CREDENTIALS=/path/to/serviceAccountKey.json
```

## Local Run

```bash
cd backend
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

Health check:

```bash
curl http://127.0.0.1:8000/health
```

## Deploy Recommendation

Use Render, Railway, Fly.io, or a VPS for this FastAPI backend.

For production, do not expose the Firebase service account file publicly. Add it only as a secure environment variable.

## Main Endpoints

```text
GET  /health
GET  /admin/settings
POST /admin/settings
POST /admin/questions/import
GET  /admin/results
```

All admin endpoints require this header:

```text
x-admin-token: YOUR_SECRET_TOKEN
```
