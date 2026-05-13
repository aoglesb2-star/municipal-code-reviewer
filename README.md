# Municipal Code Clerical Review Tool

A web app that accepts a `.docx` municipal code file, runs a non-substantive
clerical review using Claude AI, and returns a formatted Word report.

---

## What you need before starting

- A free [GitHub](https://github.com) account
- A free [Render](https://render.com) account
- An [Anthropic API key](https://console.anthropic.com) (you pay per review, ~$0.50–$2 each)

---

## Step 1 — Put the code on GitHub

1. Go to [github.com](https://github.com) and sign in.
2. Click the **+** button (top right) → **New repository**.
3. Name it `municipal-code-reviewer`. Leave everything else as default. Click **Create repository**.
4. On the next page, click **uploading an existing file**.
5. Drag ALL the files from this folder into the upload area:
   - `app.py`
   - `requirements.txt`
   - `Procfile`
   - `render.yaml`
   - `.gitignore`
   - The `templates/` folder (with `index.html` inside)
6. Scroll down, click **Commit changes**.

Your code is now on GitHub.

---

## Step 2 — Get your Anthropic API key

1. Go to [console.anthropic.com](https://console.anthropic.com) and sign in (or create an account).
2. Click **API Keys** in the left sidebar.
3. Click **Create Key**, give it a name like "municipal-reviewer", and copy the key.
4. Store it somewhere safe — you will need it in Step 3.

---

## Step 3 — Deploy on Render

1. Go to [render.com](https://render.com) and sign in.
2. Click **New +** → **Web Service**.
3. Click **Connect a repository** and authorize Render to access your GitHub.
4. Select your `municipal-code-reviewer` repository.
5. Render will auto-detect the settings from `render.yaml`. Confirm:
   - **Environment:** Python
   - **Build command:** `pip install -r requirements.txt`
   - **Start command:** `gunicorn app:app --timeout 180 --workers 1`
6. Scroll down to **Environment Variables**. Click **Add Environment Variable**:
   - Key: `ANTHROPIC_API_KEY`
   - Value: *(paste your key from Step 2)*
7. Click **Create Web Service**.

Render will build and deploy your app (takes 2–4 minutes).

---

## Step 4 — Use it

Once Render shows **Live**, click the URL it gives you (something like
`https://municipal-code-reviewer.onrender.com`).

You will see a simple page. Upload a `.docx` municipal code file, click
**Run Clerical Review**, and wait 30–90 seconds. A `.docx` report will
download automatically.

---

## Costs

- **GitHub:** Free
- **Render (free tier):** Free — note that free tier apps "sleep" after 15 minutes
  of inactivity and take ~30 seconds to wake up on first use. Upgrade to the $7/month
  Starter plan if you want it always-on.
- **Anthropic API:** Pay per use. A typical municipal code review costs roughly
  $0.50–$2.00 depending on document length.

---

## Updating the app

To make changes, edit files in your GitHub repository (click the file → pencil icon
→ edit → commit). Render will automatically redeploy within a minute.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| "ANTHROPIC_API_KEY not configured" | Check Environment Variables in Render dashboard |
| Review takes very long / times out | Document may be very large; Render free tier has a 30s HTTP timeout — upgrade to Starter |
| "Could not extract text" | Make sure the file is a real .docx, not a renamed .doc |
| App won't start | Check the Render Logs tab for error details |
