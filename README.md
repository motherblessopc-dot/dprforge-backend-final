# DPRForge Backend (Hostinger Python App)

## Files
- `server.py` — FastAPI application
- `passenger_wsgi.py` — Hostinger Passenger entry point (wraps FastAPI in WSGI)
- `requirements.txt` — Python dependencies (Hostinger-compatible)
- `runtime.txt` — Python version (3.11)
- `.env.example` — Copy to `.env` and fill in real values

## Hostinger Python App settings
- **Python version**: 3.11
- **Application startup file**: `passenger_wsgi.py`
- **Application entry point**: `application`

## Install dependencies (via SSH)
```
source ~/virtualenv/dprforge-backend/3.11/bin/activate
cd ~/dprforge-backend
pip install -r requirements.txt
```

## Environment variables — create `.env` file in this folder with:
```
MONGO_URL=mongodb+srv://USER:PASS@cluster.mongodb.net/dprforge?retryWrites=true&w=majority
DB_NAME=dprforge
CORS_ORIGINS=https://dprforge.com,https://www.dprforge.com
JWT_SECRET=long-random-string-here
ADMIN_EMAILS=owner@dprforge.com
```

See top-level `HOSTINGER-SETUP.md` for full step-by-step.
