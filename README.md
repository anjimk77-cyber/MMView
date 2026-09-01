# Shrimp FarmFlow - KMN — Marketing Manager View

A read-only Streamlit app for marketing managers, sharing the same Google
Sheet and `Customer List.xlsx` as the data-entry app and the full manager
app. It includes only:

1. 📋 Enter Customer Details (Customer / Farm selector)
2. 📊 All Saved Records
3. 🟦 Pond Layout
4. 🧾 Sales Details (with Feed Order Status boxes)
5. 🌾 All Harvest Details

It does **not** include the Zone-Wise Harvest breakdown, Running List, or
Species-wise Pond Summary sections from the full manager app.

Two actions still write to the Google Sheets (same as the manager app):
removing a date from Sales Details writes `Settle = Yes`, and removing a
row from All Harvest Details writes `Harvest Status = H`. Both persist
across refreshes. Everything else is read-only.

## Files

```
kmn-marketing-manager-app/
├── app.py                          # the Streamlit app
├── requirements.txt                # Python dependencies
├── Customer List.xlsx              # you must add this yourself (see below)
├── .gitignore
├── README.md
└── .streamlit/
    └── secrets.toml.example        # template — copy to secrets.toml locally
```

## 1. Add your Customer List file

Copy your existing `Customer List.xlsx` (the same one used by the
data-entry app) into the root of this project, next to `app.py`. It must
have at least these columns: `Customer Name`, `Farm Name with Code`,
`Zone`, `Area`. Optional but used if present: `Marketing Manager`, and
one of `Customer Code` / `Customer ID` / `Customer Code with Code` /
`Code` / `Cust Code` (used to match rows in the Sales Details sheet).

## 2. Google Sheets access

This app needs a Google Cloud **service account** with access to two
spreadsheets:

- The main `WaterQualityData` sheet (same one the data-entry and manager
  apps use)
- The Sales Details sheet (spreadsheet ID is hard-coded near the top of
  `app.py` as `SALES_SHEET_ID` — update it if your Sales Details sheet
  has a different ID)

If you already have a service account set up for the other two apps, you
can reuse the exact same credentials and just share both sheets with that
service account's email as **Viewer** (this app never needs Editor
access except for the two recycle-bin flag writes, so Editor is safer if
you want those to work).

Steps if you're starting from scratch:

1. In [Google Cloud Console](https://console.cloud.google.com/), create a
   project (or reuse the existing one) and enable the **Google Sheets
   API** and **Google Drive API**.
2. Create a **Service Account**, then create a JSON key for it and
   download it.
3. Open the main `WaterQualityData` Google Sheet and the Sales Details
   Google Sheet, and share both with the service account's `client_email`
   (found in the JSON key) — Editor access if you want the recycle-bin
   deletes to work, Viewer if this should be fully read-only (in that
   case remove or guard the two write blocks in `app.py`).

## 3. Configure secrets

Copy the template and fill in your real values:

```bash
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Open `.streamlit/secrets.toml` and paste in:
- The full contents of your service account JSON under `[gcp_service_account]`
  (each JSON key becomes a `key = "value"` line)
- Your main sheet's ID under `[gsheet] sheet_id`
- The worksheet/tab name if it isn't `WaterQualityData`

`secrets.toml` is already in `.gitignore` — never commit it.

## 4. Run locally

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## 5. Push to GitHub

```bash
cd kmn-marketing-manager-app
git init
git add .
git commit -m "Marketing manager view app"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

Double check `git status` shows `secrets.toml` is **not** staged before
you commit — only `secrets.toml.example` should go up.

## 6. Deploy on Streamlit Community Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in
   with GitHub.
2. Click **New app**, pick this repository, branch `main`, and file path
   `app.py`.
3. Before (or right after) deploying, open **Advanced settings → Secrets**
   and paste in the full contents of your local `secrets.toml`.
4. Deploy. You'll get a dedicated URL for this app — separate from the
   data-entry app and the full manager app — which you can share with
   marketing managers.

## Notes / things you may want to adjust

- **Scoping to a specific marketing manager:** right now every marketing
  manager who opens this link can pick any Customer/Farm from the full
  list — the `Marketing Manager` column (if present in `Customer
  List.xlsx`) is only shown as a caption, it isn't used to filter the
  Customer/Farm dropdowns. If you want each marketing manager to only see
  their own farms, the cleanest options are: (a) a simple login/PIN gate
  at the top of `app.py` that maps a manager to a name, then filters
  `customer_df` to rows where `Marketing Manager` matches before building
  `all_customers`; or (b) separate deployments per manager with a
  hard-coded filter. Happy to build either one if useful.
- **Recycle-bin permissions:** if marketing managers should not be able
  to permanently remove Sales Details rows or Harvest Details rows, share
  both Google Sheets with the service account as **Viewer only** — the
  two write calls will then fail gracefully with a Google API permission
  error instead of silently succeeding, though it's cleaner to just
  delete the two `st.data_editor(..., num_rows="dynamic", ...)` blocks
  and render `st.dataframe` instead if you want to remove the delete UI
  entirely.
