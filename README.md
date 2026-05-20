# 🎯 Daily Job Tracker

Automatically scrapes LinkedIn every weekday at **8:00 AM PST**, scores jobs against your profile, and updates `jobs.xlsx` in this repo.

---

## Setup (one time, ~10 minutes)

### 1. Create a private GitHub repo
Go to github.com → New repository → name it `job-tracker` → set to **Private** → Create.

### 2. Upload these files
Upload all files from this folder into the repo root:
```
job-tracker/
├── .github/workflows/daily_jobs.yml
├── scraper.py
├── requirements.txt
└── README.md
```

### 3. Add your Apify API key as a Secret
> ⚠️ Never paste your API key directly into any file — always use Secrets.

1. Go to your repo on GitHub
2. Click **Settings** → **Secrets and variables** → **Actions**
3. Click **New repository secret**
4. Name: `APIFY_API_KEY`
5. Value: paste your Apify API key (from apify.com → Settings → Integrations)
6. Click **Add secret**

### 4. Enable GitHub Actions
- Go to the **Actions** tab in your repo
- If prompted, click **"I understand my workflows, enable them"**

### 5. Run it manually the first time
- Actions tab → **Daily Job Tracker** → **Run workflow** → **Run workflow**
- Wait ~2 minutes for it to finish
- A `jobs.xlsx` file will appear in your repo — download it!

---

## Viewing your jobs

1. Go to your repo on GitHub
2. Click `jobs.xlsx`
3. Click **Download**
4. Open in Excel or Google Sheets

The sheet has:
| Column | What it shows |
|---|---|
| Match % | How well the job fits your profile (color coded) |
| Chance | High / Medium / Low / Very Low |
| Applicants | Fewer = better odds |
| Key Skills Matched | Why it scored the way it did |
| Link | Click "Apply →" to go straight to LinkedIn |
| Status | Update this yourself as you apply |
| Notes | Your personal notes per job |

---

## Schedule
Runs automatically **Monday–Friday at 8:00 AM PDT**.
New jobs are **appended** — jobs you've already seen are never duplicated.
You can also trigger it anytime manually from the Actions tab.

---

## Updating your profile / search terms
Edit `scraper.py`:
- `PROFILE["skills"]` — add/remove skills to match against
- `SEARCH_URLS` — add new LinkedIn search URLs for different roles or locations
