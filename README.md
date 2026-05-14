# DPRForge Backend — Railway Deployment

## Quick deploy to Railway

1. Create a new Railway project and connect this repo (or upload this folder).
2. Add the **MongoDB plugin** in Railway (or paste a MongoDB Atlas URI).
3. Set the environment variables from `.env.example` in the Railway "Variables" tab.
   - **CRITICAL**: `JWT_SECRET` must be a long random string (use `openssl rand -hex 32`).
   - `MONGO_URL` is auto-injected when you add Railway's MongoDB plugin.
4. Railway auto-detects the `Procfile` and runs:
   ```
   uvicorn server:app --host 0.0.0.0 --port $PORT
   ```
5. After deploy, visit `https://<your-app>.up.railway.app/api/company` — should return JSON.

## Admin login
- Auto-seeded on every boot from `ADMIN_SEED_EMAIL` + `ADMIN_SEED_PASSWORD`.
- Default: `motherblessopc@gmail.com` / `Admin1234`.
- The seed RESETS the password on every restart — change `ADMIN_SEED_PASSWORD` and redeploy to rotate.

## Endpoints quick reference
- `POST /api/auth/login` — regular user login
- `POST /api/auth/admin-login` — admin-only login (separate portal)
- `POST /api/auth/quick-buy` — guest signup (₹799 one-time)
- `GET  /api/projects/{id}/pricing` — returns 799 for guests, 599 for logged-in
- `POST /api/projects/{id}/pay-from-wallet` — debit wallet for project
- `GET  /api/projects/{id}/download/free-watermarked-pdf` — uses 1 free_dpr_credit
- `POST /api/admin/wallet/credit` — admin can credit any user's wallet
- `POST /api/admin/settings` — admin can change prices
