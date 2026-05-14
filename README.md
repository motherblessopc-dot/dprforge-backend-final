# DPRForge Backend (FastAPI + MongoDB)

## Setup
```bash
pip install -r requirements.txt
# Copy .env.example -> .env and fill values
uvicorn server:app --host 0.0.0.0 --port 8001
```

## Admin login (hidden URL)
- Frontend route: `/mb-admin-portal-7300213623`
- Email: `motherblessopc@gmail.com`
- Password: `Admin1234`

Admin is auto-seeded on every server start using `ADMIN_SEED_*` env values.

## API features
- 7-tab admin: dashboard / users (block) / payments (UTR verify) / inquiries (new/processing/completed) / DPR templates (PDF upload, price) / settings (UPI, Razorpay, WhatsApp, pricing) / analytics (date-range, exports)
- Audit log for every mutating admin action: `GET /api/admin/audit-logs`
- Excel exports: `GET /api/admin/exports/payments.xlsx`, `GET /api/admin/exports/sales.xlsx`
- Source-zip download (admin only): `GET /api/admin/source-zip/{backend|frontend|full}`
- DPR/CMA reports in Government & Bank approved format (PMEGP / Mudra / Stand-Up India / NABARD compliant, RBI Nayak working capital methodology)
