# 🎯 Daily Internship & Fellowship Tracker

Automatically pulls live internship and fellowship postings every weekday at **8:00 AM PST**, scores them against your resume/skills profile, updates `jobs.xlsx` in this repo, and posts a summary to Discord if configured.

---

## How it sources jobs (no more Apify/LinkedIn)

Instead of scraping LinkedIn (which kept getting blocked/rate-limited), this pulls **directly from the free public job-board APIs** that most startups and a lot of big tech companies use to power their own careers pages:

- **Greenhouse** (`boards-api.greenhouse.io`)
- **Lever** (`api.lever.co`)
- **Ashby** (`api.ashbyhq.com`)

These are JSON APIs meant to be publicly readable (they power the "Careers" page on each company's own site), so there's no blocking, no API key, and no cost. The tradeoff: it only sees companies that use one of these three platforms, and it needs a curated company list — that list lives at the top of `scraper.py` as `COMPANIES`, currently ~59 companies split across Seattle-area startups (Rover, Outreach, Highspot, Smartsheet, Amperity, Textio, Qumulo, PayScale, Bungie, Adaptive Biotechnologies...), remote-friendly AI/tech startups (Anthropic, OpenAI, Perplexity, Notion, Ramp, Together AI...), and larger tech companies (Stripe, Databricks, Coinbase, MongoDB...).

**To add a company:** find its careers page — if the URL looks like `jobs.lever.co/COMPANY`, `boards.greenhouse.io/COMPANY`, or `jobs.ashbyhq.com/COMPANY`, that `COMPANY` slug is the token. Add an entry to `COMPANIES` with that token and platform. The scraper skips (and logs) any token that doesn't resolve, so a bad guess never breaks the run.

## How filtering + scoring works

1. **Title filter** — only postings whose title contains `intern`, `co-op`, `new grad`, or `early career` survive at all (drops the senior/staff SWE postings that are also on these boards).
2. **Function filter** — drops recruiting/HR/sales/marketing postings that happen to say "early career" or "intern" but aren't engineering roles.
3. **Hard seniority / grad-level filter** — titles containing `senior`, `staff`, `principal`, `director`, `manager`, `PhD`, `research scientist`, `master's student`, `doctoral`, `graduate researcher`, or `MBA` are dropped outright, no matter how many skills match. **Only undergraduate-eligible roles survive** (internship / co-op / fellowship / new-grad) — nothing requiring an in-progress or completed graduate degree.
4. **Domain exclusion** — cybersecurity (security engineer/analyst, infosec, pentesting, red/blue team), aerospace (aerospace, avionics, satellite, spacecraft, propulsion, flight software), and non-CS functions (finance, legal, accounting, medical/clinical, HR, sales, marketing) are dropped, checked against both title *and* description so a generic "Software Engineer Intern" on an aerospace team, or a "Medical Fellow" track sharing boilerplate text with an ML fellowship, still gets excluded.
5. **Skill requirement** — needs at least 2 real skill/keyword overlaps with your profile; a single generic hit (e.g. an accounting internship mentioning "Excel") isn't enough.
6. **Scoring (0-100)** for everything that survives:
   - Skill/keyword overlap with your profile (including AI/data-analytics/business-analyst/ML terms) — 0-35 pts
   - Internship > fellowship > co-op > new-grad title strength — 0-25 pts
   - Seattle/WA or remote location — 0-20 pts
   - Recency (posted <7/14/30 days ago) — 0-20 pts
   - Penalties for "5+ years experience", "master's degree required", "security clearance", or any PhD mention — -15 pts each
6. Only postings scoring **40+** get written to the sheet at all — the point is a short, high-confidence list, not everything ranked.

## Setup (one time, ~5 minutes)

### 1. Push these files to your existing `job-tracker` repo
Replace `scraper.py`, `requirements.txt`, `.github/workflows/daily_jobs.yml`, and this `README.md` with the versions in this folder.

### 2. Remove the old Apify secret (optional)
The scraper no longer uses Apify, so you can delete the `APIFY_API_KEY` secret from **Settings → Secrets and variables → Actions** if you want — nothing reads it anymore. Every job source here is public/no-auth.

### 3. Add your Discord webhook (optional but recommended)
To get a daily Discord message when new matches are found:
1. Repo → **Settings → Secrets and variables → Actions → New repository secret**
2. Name: `DISCORD_WEBHOOK_URL`
3. Value: your Discord channel's webhook URL (Discord: channel **Settings → Integrations → Webhooks → New Webhook → Copy URL**)

If this secret isn't set, the scraper just skips the notification — the Excel sheet still updates normally.

### 4. Run it manually the first time
- Actions tab → **Daily Job Tracker** → **Run workflow** → **Run workflow**
- Wait ~1-2 minutes
- `jobs.xlsx` will appear in your repo

---

## Viewing your jobs

Download `jobs.xlsx` from the repo and open in Excel/Google Sheets. Columns:

| Column | What it shows |
|---|---|
| Match % | Fit score, color coded |
| Fit | Strong Fit (70+) / Good Fit (55+) / Possible Fit (40+) |
| Source | Which ATS it came from (gh / lever / ashby) |
| Why It Matched | The specific skills/signals that drove the score |
| Link | Click "Apply →" to go straight to the posting |
| Status | Update yourself as you apply |
| Notes | Your own notes per job |

The current `jobs.xlsx` (every match found to date) is also attached directly to each Discord message, so you don't need to open GitHub to see the full list — just download the attachment from Discord.

---

## Schedule
Runs automatically **Monday–Friday at 8:00 AM PDT**. New postings are **appended**; nothing already in the sheet is duplicated. You can also trigger it anytime from the Actions tab.

## Tuning it further
All of this lives at the top of `scraper.py`:
- `PROFILE["skills"]` — add/remove keywords as your stack changes (e.g. once your RAG/stock-analyzer work is resume-ready, keywords like `rag`, `langchain`, `embeddings` are already in here)
- `PROFILE["preferred_locations"]` — currently Seattle-area + remote only; add other states back if you widen the search
- `COMPANIES` — add more companies as you find their board tokens
- `MIN_SCORE_TO_INCLUDE` — raise it (e.g. to 55) if 40 is still surfacing too much; lower it if the list feels too thin
- `INTERNSHIP_TITLE_RE` — loosen this if you also want full-time junior/new-grad SWE roles, not just internships
