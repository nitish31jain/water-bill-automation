# 💧 Water Bill — Per-Litre Cost Calculator

A small [Streamlit](https://streamlit.io/) app for a residential society. It computes the
**per-litre water cost for the previous month** from the society's consumption report and the
three water-supply bills, so residents can be charged by actual usage.

## What it does

1. Upload the monthly **Consumption Report** (`.xlsx`) — the filename must contain the month and
   year, e.g. `Consumption Report May-2026.xlsx`.
2. Upload the **Reimbursement template** (`.csv`).
3. Enter the previous-month supply bills (₹): **Cauvery**, **SSE tanker**, **CKE tanker**.
4. The app validates that the sheet is for the **previous month** (derived from today's date) and
   computes the per-litre cost.

### Billing rules

- The bill can only ever be generated for the **previous month**. Current, future, or older
  months are rejected — current-month generation is explicitly blocked.
- Each flat's reading comes from the consumption sheet's **Total** column.
- Merged meters (e.g. `C 1303 & 1403`) are consolidated to a single flat row (`C-1303`).
- Any flat whose total is **below 50 litres** is billed as **0**.
- **Per-litre cost = (Cauvery + SSE + CKE) ÷ grand total of the updated reimbursement sheet.**

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the URL shown (usually http://localhost:8501).

## Deploy (Streamlit Community Cloud)

1. Push this repo to GitHub (already done if you're reading this on GitHub).
2. Go to https://share.streamlit.io and sign in with GitHub.
3. **New app** → pick this repository, branch `main`, main file `app.py` → **Deploy**.

No secrets or environment variables are required.

## Privacy

Resident data files (`*.xlsx`, `*.csv`) contain personal information and are **git-ignored** —
they are never committed. Each user uploads their own sheets at runtime; nothing is stored.
