# ICICI Transaction Analyser

Upload an ICICI bank Statement of Account (`.xlsx` export) and get a breakdown
of **who paid how much, and to whom** — money in by payer, money out by
payee, and a searchable/sortable view of every transaction.

## How it works

- The frontend (`static/index.html`) lets you upload a `.xlsx` file.
- The Flask backend (`app.py`) parses it with `parser.py`, which reads the
  bank's free-text `Narration` field on each row and extracts a best-effort
  counterparty name and transaction category (UPI, IMPS, NEFT, RTGS, cheque
  clearing, bill payment, etc.).
- The parsed transactions and aggregated summary are returned as JSON and
  rendered in the browser — nothing is written to disk or stored server-side.

### Expected file format

The uploaded `.xlsx` must have a header row with (at least) these columns,
matching ICICI's standard SOA export:

```
Ac_No, AC_Name, Tran_ID, Tran_Date, Inst_Type, Inst_Num, Dr_Amt, Cr_Amt, Balance, Narration, pstd_dt
```

If a file doesn't match this format, the app returns a clear error instead
of guessing.

## Local development

```bash
python -m venv .venv
source .venv/bin/activate       # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Visit `http://localhost:5000` and upload a statement.

## Deploying on Render

This repo includes a `render.yaml` (Render's "Blueprint" format) and a
`Procfile` as a fallback.

1. Push this repo to GitHub (already done if you're reading this on GitHub).
2. In Render: **New → Blueprint**, connect this repo, and Render will pick
   up `render.yaml` automatically (build: `pip install -r requirements.txt`,
   start: `gunicorn app:app`).
   - Alternatively, **New → Web Service**, connect the repo, set:
     - Build Command: `pip install -r requirements.txt`
     - Start Command: `gunicorn app:app`
3. Deploy. Render assigns a public URL — the app is stateless, so no
   database or persistent disk is needed.

## Notes & limitations

- ICICI truncates the `Narration` field to 50 characters, so some
  counterparty names are cut short and may represent the same entity under
  slightly different labels (e.g. `SAIRAJCON` vs `SAIRAJCONS`).
- The "Payers" and "Payees" views exclude bank charges, cheque returns, and
  unnamed cash deposits, so they reflect genuine counterparties only — the
  "All Transactions" tab still shows every row, including those.
- No statement data, PAN/Aadhaar numbers, or any other uploaded file is
  persisted anywhere by this app — uploads are parsed in-memory per request.
- The narration formats handled (`UPI`, `MMT/IMPS`, `INF/INFT`, `CLG`,
  `TRF`, `NEFT`, `RTGS`, `BIL`, `TRFR TO/FROM`, `CAM`, `CASH PAID`, bank
  charge lines) cover ICICI's common patterns; anything else falls back to
  an "Other/Unclassified" category with the raw narration as the label.
