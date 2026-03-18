# BLUCOR Daily Activity Report — Setup Guide

## Architecture
- **Frontend**: Static HTML page hosted on GitHub Pages (free, always on)
- **Backend**: Google Apps Script web app (free, serverless)
- **Database**: Google Sheets (auto-populated, easy to review)

Total cost: **$0/month**

---

## Step 1: Create the Google Sheet

1. Go to [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Name it **"Blucor Daily Activity Reports"**
3. Note the spreadsheet URL — you'll need it in Step 2

---

## Step 2: Deploy the Google Apps Script

1. In your new spreadsheet, go to **Extensions → Apps Script**
2. Delete any existing code in `Code.gs`
3. Copy the entire contents of `apps-script.gs` from this folder and paste it in
4. Click **Deploy → New deployment**
5. Click the gear icon next to "Select type" and choose **Web app**
6. Configure:
   - **Description**: "Blucor Daily Report Receiver"
   - **Execute as**: Me (your Google account)
   - **Who has access**: **Anyone**
7. Click **Deploy**
8. **Authorize** when prompted (review permissions, click Allow)
9. **Copy the Web app URL** — it looks like:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```

---

## Step 3: Configure the Form

1. Open `index.html` in a text editor
2. Find the line near the top of the `<script>` section:
   ```js
   var SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE';
   ```
3. Replace `YOUR_APPS_SCRIPT_URL_HERE` with the URL from Step 2
4. Save the file

---

## Step 4: Host on GitHub Pages

### Option A — Create a new repo (recommended)

1. Go to [github.com/new](https://github.com/new)
2. Name the repo `blucor-daily-report`
3. Make it **Public** (required for free GitHub Pages)
4. Create the repo
5. Push `index.html`:
   ```bash
   cd blucor-daily-report
   git init
   git add index.html
   git commit -m "Daily activity report form"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/blucor-daily-report.git
   git push -u origin main
   ```
6. Go to repo **Settings → Pages**
7. Source: **Deploy from a branch**
8. Branch: **main**, folder: **/ (root)**
9. Click Save

Your form will be live at:
```
https://YOUR_USERNAME.github.io/blucor-daily-report/
```

### Option B — Use an existing GitHub Pages site

Just copy `index.html` into your existing Pages repo and link to it.

---

## How It Works

1. Employee opens the link on any device (phone, tablet, computer)
2. Fills out the form
3. Clicks **Submit Report**
4. Data is sent to Google Apps Script → written as a new row in the Google Sheet
5. Form shows a success confirmation
6. They can also print a filled or blank version

---

## Google Sheet Columns

The script auto-creates a "Daily Reports" tab with these columns:

| Column | Description |
|--------|-------------|
| Timestamp | Auto-generated server time |
| Report Date | The date on the form |
| Employee Name | Who submitted |
| Role/Team | Their role |
| Tasks (summary) | All task descriptions + hours |
| Total Hours | Sum of task hours |
| Work Orders (summary) | WO numbers + locations |
| Hours Worked | From crew/safety section |
| Safety Incidents | Count |
| Safety Details | Description |
| PPE Check | Done/Issues |
| Calls/Meetings | Count |
| Change Orders | Count |
| Info Requests | Count |
| Client Details | Free text |
| Challenges | Bullet points |
| Tomorrow Focus | Numbered list |
| ... | All self-assessment fields |
| Full JSON | Complete raw data backup |

---

## Updating the Form

After updating `index.html`, just push to GitHub and the page updates within minutes:
```bash
git add index.html
git commit -m "Update report form"
git push
```

## Updating the Script

If you modify `apps-script.gs`, paste the new code into the Apps Script editor and **create a new deployment** (Deploy → Manage deployments → Edit → New version).
