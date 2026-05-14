# DPRForge Backend — Railway Deployment

## Quick deploy

1. Upload this folder to your Railway service (or connect via GitHub).
2. Add the **MongoDB plugin** (or paste a MongoDB Atlas URI as `MONGO_URL`).
3. In Railway → **Variables**, set everything from `.env.example`:
   - `MONGO_URL` (auto-injected if you used Railway's MongoDB plugin)
   - `DB_NAME=dprforge_db`
   - `JWT_SECRET` — generate with `openssl rand -hex 32`
   - `ADMIN_EMAILS=motherblessopc@gmail.com`
   - `ADMIN_SEED_EMAIL=motherblessopc@gmail.com`
   - `ADMIN_SEED_PASSWORD=Admin1234`
   - `CORS_ORIGINS=https://www.dprforge.com,https://dprforge.com`
   - `EMERGENT_LLM_KEY=<your-key>` (only needed if you re-enable AI narrative)
4. Railway auto-detects `Procfile` + `nixpacks.toml` + `runtime.txt` and starts:
   ```
   uvicorn server:app --host 0.0.0.0 --port $PORT
   ```
5. Test: `curl https://<your-app>.up.railway.app/api/company` should return JSON.

## Troubleshooting build failures

### "Failed to build image" / pip install errors
- This `requirements.txt` is the **minimal runtime-only** list. If you replace it with a
  bigger one (containing `pandas`, `numpy`, `mypy`, etc.) the build will take 5+ minutes
  and may run out of memory.
- The `emergentintegrations` package is **commented out** because it's on a custom pip
  index and can fail on Railway. AI narrative will be disabled until you re-enable it;
  every other feature (admin, wallet, payments, PDF, Excel) works regardless.

### How to re-enable AI narrative
1. Uncomment the `emergentintegrations>=0.1.0` line in `requirements.txt`.
2. If Railway still can't find it, add this to `nixpacks.toml` install command:
   ```
   pip install --extra-index-url https://d33sy5i8bnduwe.cloudfront.net/simple/ emergentintegrations
   ```
3. Set `EMERGENT_LLM_KEY` in Railway → Variables.

## Admin login
- Auto-seeded on every boot from `ADMIN_SEED_EMAIL` + `ADMIN_SEED_PASSWORD`.
- Default: `motherblessopc@gmail.com` / `Admin1234`.
- **The seed RESETS the password on every restart** — change `ADMIN_SEED_PASSWORD` and
  redeploy to rotate. Google sign-in for this same email also works (auto-promotion).

## Endpoints quick reference
- `POST /api/auth/login` — regular user login
- `POST /api/auth/admin-login` — admin-only login (separate portal)
- `POST /api/auth/quick-buy` — guest signup (₹799 one-time)
- `POST /api/auth/google-session` — Google OAuth session exchange
- `GET  /api/projects/{id}/pricing` — returns 799 for guests, 599 for logged-in
- `POST /api/projects/{id}/pay-from-wallet` — debit wallet for project
- `GET  /api/projects/{id}/download/free-watermarked-pdf` — uses 1 free_dpr_credit
- `POST /api/admin/wallet/credit` — admin can credit any user's wallet
- `POST /api/admin/settings` — admin can change prices
