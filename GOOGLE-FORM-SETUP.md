# BLUCOR Daily Activity Report — Google Form + KPI Dashboard Setup

## Overview

This system replaces the manual HTML submit button with a **Google Form** that feeds directly into a **Google Sheet** with an automated **KPI dashboard**. The HTML form at [index.html](index.html) is preserved for printing blank or filled reports.

| Component | Purpose |
|-----------|---------|
| **Google Form** | Employees fill out daily — standardized, mobile-friendly |
| **Google Sheet** | Stores all responses automatically |
| **KPI Engine** | Calculates 7 metrics per employee daily |
| **Dashboard** | Color-coded management scorecard |
| **HTML Form** | Print-only — blank forms for the field |

---

## Step 1: Create the Google Form

1. Go to [script.google.com](https://script.google.com)
2. Click **New project**
3. Name it: `BLUCOR Form Generator`
4. Delete any existing code in `Code.gs`
5. Open [form-generator.gs](form-generator.gs) and **copy the entire file contents**
6. Paste into the Apps Script editor
7. Click **Run** → select `createBluCorDailyForm`
8. Authorize when prompted (you'll see a "This app isn't verified" screen — click **Advanced → Go to BLUCOR Form Generator**)
9. Check the **Execution Log** (View → Execution log) — you'll see:
   ```
   Form URL (edit):    https://docs.google.com/forms/d/xxxxx/edit
   Form URL (submit):  https://docs.google.com/forms/d/xxxxx/viewform
   Linked Sheet ID:    xxxxxxxxxxxxxxx
   Sheet URL:          https://docs.google.com/spreadsheets/d/xxxxx
   ```
10. **Bookmark the Form submit URL** — this is what you'll share with employees
11. **Bookmark the Sheet URL** — this is where responses and the dashboard will live

### Test the Form
1. Open the Form submit URL
2. Fill out a test response (use your own name)
3. Submit
4. Open the Sheet URL — you should see a `Form Responses 1` tab with the data

---

## Step 2: Set Up the KPI Dashboard

1. Open the **linked Google Sheet** (from Step 1)
2. Go to **Extensions → Apps Script**
3. This opens the Sheet's script editor
4. Delete any existing code
5. Open [dashboard-setup.gs](dashboard-setup.gs) and **copy the entire file contents**
6. Paste into the Apps Script editor
7. Click **Run** → select `setupDashboard`
8. Authorize when prompted

### What gets created:

| Tab | Contents |
|-----|----------|
| **Config** | Settings: expected work days/week, color thresholds |
| **KPI Engine** | Per-employee metric calculations (20 columns) |
| **Dashboard** | Management scorecard with color-coded scores |

### Auto-Refresh
A daily trigger is automatically created — `updateKPIs()` runs every day at 5 AM to recalculate all metrics. To refresh manually, run `updateKPIs()` from the script editor.

---

## Step 3: Share with Employees

1. Open the Google Form (edit URL)
2. Click the **Send** button (paper airplane icon)
3. Options:
   - **Email** — send directly to employees
   - **Link** — copy the short URL and share via text/chat
   - **Embed** — if you want to embed in a company intranet page

### Recommended: Create a Bookmark
Share this URL format with employees:
```
https://docs.google.com/forms/d/YOUR_FORM_ID/viewform
```

---

## Step 4: Update the HTML Form (Optional)

The HTML form at [index.html](index.html) has been updated to include a banner linking to the Google Form. The submit button has been removed since submissions now go through Google Forms.

The HTML form is still useful for:
- **Printing blank reports** for field workers without phones
- **Printing filled reports** for physical records

---

## KPI Reference

| KPI | Calculation | Weight |
|-----|------------|--------|
| **Productivity Score** | AVG(effectiveness rating, time mgmt rating) / 10 × 100 | 25% |
| **Safety Compliance** | (PPE "Done" rate + zero-incident rate) / 2 | 20% |
| **Initiative Score** | (Extra tasks % + proactive work % + improvements %) / 3 | 15% |
| **Quality Score** | (Error check % + normalized check count) / 2 | 15% |
| **Submission Consistency** | Actual submissions / expected work days × 100 | 15% |
| **Mentorship Rate** | Days mentored / total submissions × 100 | 10% |
| **Communication Volume** | AVG daily internal + external contacts (not scored, informational) | — |

### Color Thresholds (configurable in Config tab)
| Color | Score |
|-------|-------|
| 🟢 Green | ≥ 80% |
| 🟡 Yellow | 50–79% |
| 🔴 Red | < 50% |

### Overall Grade
Weighted average of all scored KPIs (weights shown above).

---

## Troubleshooting

### "No Form Responses sheet found"
The dashboard script looks for a tab named `Form Responses 1` (or any tab starting with "Form Responses"). Make sure:
- The Form is linked to this Sheet (Step 1)
- At least one test response has been submitted

### KPIs show 0 for everything
Submit at least 2–3 test responses with different employee names, then run `updateKPIs()` manually.

### Column mapping errors
If Google changes the Form question titles (or you edit them), the dashboard may not find the right columns. The `mapColumns()` function matches by header text — run `updateKPIs()` after any Form edits and check the Execution Log for warnings.

### Want to change thresholds?
Edit the **Config** tab → cells B3 (green) and B4 (yellow). Changes take effect on the next `updateKPIs()` run.

---

## File Inventory

| File | Purpose |
|------|---------|
| [index.html](index.html) | Printable HTML form (not for submission) |
| [form-generator.gs](form-generator.gs) | Creates the Google Form programmatically |
| [dashboard-setup.gs](dashboard-setup.gs) | Creates KPI Engine + Dashboard tabs |
| [GOOGLE-FORM-SETUP.md](GOOGLE-FORM-SETUP.md) | This file — setup instructions |
| [SETUP.md](SETUP.md) | Legacy setup guide (original HTML submit flow) |
