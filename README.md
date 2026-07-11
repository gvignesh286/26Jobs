# 🎯 Daily Internship & Fellowship Tracker

Automatically pulls live internship and fellowship postings every weekday at **8:00 AM PST**, scores them against your resume/skills profile, updates `jobs.xlsx` in this repo, and posts a summary to Discord if configured.

---

## How it sources jobs (no more Apify/LinkedIn)

Instead of scraping LinkedIn (which kept getting blocked/rate-limited), this pulls **directly from the free public job-board APIs** that most startups and a lot of big tech companies use to power their own careers pages:

- **Greenhouse** (`boards-api.greenhouse.io`)
- **Lever** (`api.lever.co`)
- **Ashby** (`api.ashbyhq.com`)

These are JSON APIs meant to be publicly readable (they power the "Careers" page on each company's own site), so there's no blocking, no API key, and no cost. The tradeoff: it only sees companies that use one of these three platforms, and it needs a curated company list — that list lives at the top of `scraper.py` as `COMPANIES`, currently **126 companies** split across Seattle-area startups (Rover, Outreach, Highspot, Smartsheet, Amperity, Textio, Qumulo, PayScale, Bungie, Adaptive Biotechnologies...), remote-friendly AI/tech startups (Anthropic, OpenAI, Perplexity, Notion, Ramp, Together AI, Cohere, LangChain...), and larger/public tech companies (Stripe, Databricks, Coinbase, MongoDB, Palantir, Snowflake, Cloudflare...).

The company list got expanded from 59 → 126 because most of the daily "new" postings were the same handful of open reqs getting deduped day after day — a bigger candidate pool is what actually drives more genuinely-new matches per day, not a lower score bar. If volume drops off again later (e.g. once summer 2027 recruiting season winds down), add more companies rather than lowering `MIN_SCORE_TO_INCLUDE`.

**To add a company:** find its careers page — if the URL looks like `jobs.lever.co/COMPANY`, `boards.greenhouse.io/COMPANY`, or `jobs.ashbyhq.com/COMPANY`, that `COMPANY` slug is the token. Add an entry to `COMPANIES` with that token and platform. The scraper skips (and logs) any token that doesn't resolve, so a bad guess never breaks the run.

## How filtering + scoring works

1. **Title filter** — only postings whose title contains `intern`, `co-op`, `new grad`, or `early career` survive at all (drops the senior/staff SWE postings that are also on these boards).
2. **Function filter** — drops recruiting/HR/sales/marketing postings that happen to say "early career" or "intern" but aren't engineering roles.
3. **Hard seniority / grad-level filter** — titles containing `senior`, `staff`, `principal`, `director`, `manager`, `PhD`, `research scientist`, `master's student`, `doctoral`, `graduate researcher`, or `MBA` are dropped outright, no matter how many skills match. **Only undergraduate-eligible roles survive** (internship / co-op / fellowship / new-grad) — nothing requiring an in-progress or completed graduate degree.
4. **Domain exclusion** — cybersecurity (security engineer/analyst, infosec, pentesting, red/blue team), aerospace (aerospace, avionics, satellite, spacecraft, propulsion, flight software), and non-CS functions (finance, legal, accounting, medical/clinical, HR, sales, marketing) are dropped, checked against both title *and* description so a generic "Software Engineer Intern" on an aerospace team, or a "Medical Fellow" track sharing boilerplate text with an ML fellowship, still gets excluded.
5. **Skill requirement** — needs at least 2 real skill/keyword overlaps with your profile; a single generic hit (e.g. an accounting internship mentioning "Excel") isn't enough.
6. **US-only** — any posting whose location names a foreign country/city (London, Toronto, Bangalore, Sydney, etc.) is dropped outright, regardless of score.
7. **Scoring (0-100)** for everything that survives:
   - Skill/keyword overlap with your profile (including AI/data-analytics/business-analyst/ML terms) — 0-35 pts
   - Internship > fellowship > co-op > new-grad title strength — 0-25 pts
   - Seattle/WA or remote location — 0-20 pts
   - Recency (posted <7/14/30 days ago) — 0-20 pts
   - Penalties for "5+ years experience", "master's degree required", "security clearance", or any PhD mention — -15 pts each
8. Only postings scoring **50+** get written to the sheet at all — the point is a short, high-confidence list, not everything ranked.

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

`jobs.xlsx` has two tabs:
- **Internship Matches** — every qualifying posting
- **Startups** — narrowed to companies tagged `startup: True` in `COMPANIES` **and** located in Seattle, the SF Bay Area, Texas, Arizona, or remote (`STARTUP_TAB_LOCATION_RE` at the top of `scraper.py`). Public/large-cap companies (Palantir, Snowflake, Stripe...) and startup postings outside those regions stay in the main tab only.

Download `jobs.xlsx` from the repo and open in Excel/Google Sheets. Columns (same on both tabs):

| Column | What it shows |
|---|---|
| Company Website | Click "Website →" — the company's homepage |
| Match % | Fit score, color coded |
| Fit | Strong Fit (70+) / Good Fit (55+) / Possible Fit (50+) |
| Source | Which ATS it came from (gh / lever / ashby) |
| Why It Matched | The specific skills/signals that drove the score |
| Link | Click "Apply →" to go straight to the posting |
| WSU Alumni Search | Click "WSU Alumni →" — opens a LinkedIn people search for `{Company} Washington State University`, logged in as you |
| Recruiter Search | Click "Recruiters →" — opens a LinkedIn people search for `{Company} university recruiting` |
| LinkedIn Note Draft | A ready-to-send connection note (under 300 chars), tailored to the role — pick a person from the search above and paste this in |
| Email Draft | A longer version for email/InMail, with a subject line and your contact info already filled in |
| Status | Update yourself as you apply |
| Notes | Your own notes per job |

The current `jobs.xlsx` (every match found to date) is also attached directly to each Discord message, so you don't need to open GitHub to see the full list — just download the attachment from Discord.

### On the outreach columns — what this does and doesn't do
The search-link and message-draft columns are meant to speed up warm outreach, not automate it:
- **No LinkedIn scraping.** The two search links just open LinkedIn's own people-search with useful filters pre-filled — you still browse the results and pick a real person yourself, logged into your own account. Automating LinkedIn scraping violates their ToS and risks your account getting flagged.
- **No auto-sending.** The message drafts are text sitting in a cell. Nothing sends anything on your behalf — you copy the draft, personalize it for whoever you found, and send it yourself from LinkedIn/email.
- The message draft picks its "why I'd be a good fit" line based on which of your skills matched that specific posting (ML-heavy roles get the APEX Stock Scanner pitch, web-heavy roles get TravelBuddy/the hackathon project, data-heavy roles get the stock-analysis/trail-mapping projects) — tune the wording in `HIGHLIGHTS` at the top of `scraper.py` if it doesn't sound like you.

---

## Schedule
Runs automatically **Monday–Friday at 8:00 AM PDT**. New postings are **appended**; nothing already in the sheet is duplicated. You can also trigger it anytime from the Actions tab.

## Tuning it further
All of this lives at the top of `scraper.py`:
- `PROFILE["skills"]` — add/remove keywords as your stack changes (e.g. once your RAG/stock-analyzer work is resume-ready, keywords like `rag`, `langchain`, `embeddings` are already in here)
- `PROFILE["preferred_locations"]` — currently Seattle-area + remote only; add other states back if you widen the search
- `COMPANIES` — add more companies as you find their board tokens
- `MIN_SCORE_TO_INCLUDE` — currently 50; raise it further (e.g. to 60) if that's still surfacing too much, or lower it if the list feels too thin
- `NON_US_LOCATION_RE` — the sheet is US-only right now; remove a country/city from this pattern (or drop the check in `score_job` entirely) if you want to reopen it to a specific country
- `INTERNSHIP_TITLE_RE` — loosen this if you also want full-time junior/new-grad SWE roles, not just internships
