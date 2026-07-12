# 🎯 Daily Internship & Fellowship Tracker

Automatically pulls live internship and fellowship postings every weekday at **8:00 AM PST**, scores them against your resume/skills profile, updates `jobs.xlsx` in this repo, and posts a summary to Discord if configured.

---

## How it sources jobs (no more Apify/LinkedIn)

Instead of scraping LinkedIn (which kept getting blocked/rate-limited), this pulls **directly from the free public job-board APIs** that most startups and a lot of big tech companies use to power their own careers pages:

- **Greenhouse** (`boards-api.greenhouse.io`)
- **Lever** (`api.lever.co`)
- **Ashby** (`api.ashbyhq.com`)

These are JSON APIs meant to be publicly readable (they power the "Careers" page on each company's own site), so there's no blocking, no API key, and no cost. The tradeoff: it only sees companies that use one of these three platforms, and it needs a curated company list — that list lives at the top of `scraper.py` as `COMPANIES`, currently **135 companies** split across Seattle-area startups (Rover, Outreach, Highspot, Smartsheet, Amperity, Textio, Qumulo, PayScale, Bungie, Adaptive Biotechnologies...), remote-friendly AI/tech startups (Anthropic, OpenAI, Perplexity, Notion, Ramp, Together AI, Cohere, LangChain...), larger/public tech companies (Stripe, Databricks, Coinbase, MongoDB, Palantir, Snowflake, Cloudflare...), and Texas/Arizona companies (Carvana, Axon, CS Disco, Bazaarvoice, Homeward, ShiftKey, AlertMedia, Convey, Shipwell).

The company list got expanded from 59 → 126 because most of the daily "new" postings were the same handful of open reqs getting deduped day after day — a bigger candidate pool is what actually drives more genuinely-new matches per day, not a lower score bar. If volume drops off again later (e.g. once summer 2027 recruiting season winds down), add more companies rather than lowering `MIN_SCORE_TO_INCLUDE`.

**On regional coverage (Seattle / TX / AZ):** if the Startups tab looks thin for a specific region on a given day, that's very likely real — companies don't have open internship reqs every day, especially smaller startups that don't run formal internship programs at all. Before assuming it's a bug, run the diagnostic below to check whether qualifying postings actually exist for that region right now:
```python
from scraper import scrape_jobs
jobs = scrape_jobs()
for j in jobs:
    print(j['companyName'], '|', j['location'])
```
126 → 135 added Carvana/Axon/CS Disco/Bazaarvoice (Tempe/Scottsdale/Austin — public companies, main tab only) and Homeward/ShiftKey/AlertMedia/Convey/Shipwell (Austin/Dallas — private, would show on Startups tab) specifically to close a real gap: the list previously had zero companies with any TX/AZ presence at all.

**To add a company:** find its careers page — if the URL looks like `jobs.lever.co/COMPANY`, `boards.greenhouse.io/COMPANY`, or `jobs.ashbyhq.com/COMPANY`, that `COMPANY` slug is the token. Add an entry to `COMPANIES` with that token and platform. The scraper skips (and logs) any token that doesn't resolve, so a bad guess never breaks the run.

### Layered source: SimplifyJobs aggregator
On top of the 135 hand-picked companies, `fetch_simplify()` pulls from [SimplifyJobs/Summer2026-Internships](https://github.com/SimplifyJobs/Summer2026-Internships)'s `listings.json` — a community-maintained, bot-updated (every 30-60 min) file aggregating postings from thousands of companies across many ATS platforms. This is what actually fixed the "only a couple new postings a day" problem — it found real Seattle (TikTok, ByteDance, Amazon, Docugami, OfferUp) and Texas (Copart, Optiver, Citadel, Exowatt) postings that the 135-company list had no way to reach.

Tradeoffs vs. the direct ATS sources:
- **No description text** — the aggregator only gives title/company/location/category, so `CATEGORY_TEXT` substitutes generic category-based text (e.g. "software engineering full stack...") for scoring purposes. This means Simplify-sourced postings in the same category tend to land on similar scores — less differentiated than ATS-sourced postings, which get scored against real job descriptions.
- **Degree filtering happens at the source** — the JSON has a `degrees` field per posting; we only keep entries listing `"Bachelor's"`, same undergrad-only bar as everywhere else.
- **Company-level defense blocklist** (`DEFENSE_COMPANY_BLOCKLIST`) — without description text, the usual `DOMAIN_EXCLUDE_RE` keyword check can miss defense/military-hardware companies whose internship titles don't literally say "aerospace" (e.g. Anduril, Saronic). Add a company name here if another one slips through.
- **Cross-source dedup** — if a company is tracked both directly (via `COMPANIES`) and shows up in the aggregator for the same role, the direct-ATS copy wins (real description text beats the synthetic one), and the Simplify duplicate is dropped by company+title match.
- **It's community-scraped, not an official API** — if the repo's schema or URL ever changes, `fetch_simplify()` will need updating; it fails gracefully (prints an error, doesn't break the other 135 sources) if the fetch fails.

### Layered source: Y Combinator's Work at a Startup
`fetch_yc()` pulls from `workatastartup.com/jobs/l/{software-engineer,product-manager,science}` — the public job listing pages (`robots.txt` allows it, no login needed to view) for [Y Combinator](https://www.ycombinator.com/)-backed startups. No official API: this parses the `data-page` JSON payload the page's Inertia.js frontend embeds server-side. Every listing here is by definition a YC-backed startup, so these always feed the Startups tab.

We looked at Dice too — **not built**. Dice's `robots.txt` explicitly disallows the job-search paths (`/jobs?q*`, `/job`, `/jobsearch/`), which is Dice stating they don't want automated access to listings. Same category of thing as the LinkedIn scraping this project already avoids.

Tradeoffs / things to know about the YC source specifically:
- **Covers both internships and full-time roles** — unlike every other source, `scrape_jobs()` does *not* require `INTERNSHIP_TITLE_RE` to match for YC postings, since small startups rarely label roles "new grad" the way big companies do. Filtering instead relies entirely on `HARD_AVOID_TITLE_RE` (which now also excludes `founding` — as in "Founding Engineer"/"Founding Solutions Engineer" — and `head of`, on top of the usual senior/staff/lead/manager/CTO exclusions) and `NON_TECHNICAL_ROLE_RE`.
- **Same 55-point bar as every other source, no exception** — YC postings have no posting date (0 of the 0-20 recency points are ever achievable) and thinner synthetic description text (`YC_ROLE_TYPE_TEXT` + the company's one-liner, in place of a real job description) than the other sources. To compensate without lowering the bar itself, the YC full-time title-tier credit is boosted to 23 (vs. 15 originally) — genuinely strong matches (good skill overlap + a preferred location) can still clear 55; weaker ones correctly get filtered out like everywhere else.
- **Defense-company screening uses the company's one-liner, not just a name blocklist** — YC gives us `companyOneLiner` (e.g. Hop Aero's "Rocket cargo delivery to contested environments"), so `fetch_yc()` runs that text through `DOMAIN_EXCLUDE_RE` in addition to checking `DEFENSE_COMPANY_BLOCKLIST` by name.
- **More fragile than the JSON-API sources** — if YC changes their frontend framework or page markup, `fetch_yc()`'s regex extraction may break and need updating. It fails gracefully per category URL (prints an error, the other sources keep working).

### Startup-tag gap for aggregator-discovered companies
The Startups tab only shows postings where `isStartup` is `True`. For the 135 curated `COMPANIES`, that tag is set by hand. For companies discovered only through SimplifyJobs or YC, there's no such data — so `isStartup` defaults to `False` unless the company is YC-sourced (where it's always `True` by definition) or listed in `ADDITIONAL_STARTUP_COMPANIES`. This caused a real bug: Seattle/TX startups found via SimplifyJobs (Docugami in Kirkland, Exowatt in Austin) were invisible on the Startups tab even though they qualified, while only YC entries (which cluster in SF) ever showed up — making the tab look SF-only. Fixed by adding known-good companies to `ADDITIONAL_STARTUP_COMPANIES` as they're spotted; it's a manually-grown list, not automatic classification, so add to it whenever a real non-SF startup gets missed.

### University career pages (UW, UT Austin, UC system) — investigated, not built
These are all gated behind licensed third-party platforms tied to institutional login, not something scrapable:
- **UW Seattle**: their career center states outright that Handshake is UW's job/internship database — everything lives there, gated by UW NetID SSO.
- **UT Austin**: uses a different platform (12twenty@Texas), same pattern — gated by UT EID SSO.
- **UC system**: near-certain to follow the same pattern — Handshake alone serves 1,400+ universities.

The blocker isn't robots.txt — it's that employers choose which specific schools can see their postings, and only that school's own students can log in. A WSU student has no UW/UT credentials to use, and even with them, scraping a login-gated platform is the same category of thing already declined for LinkedIn and Dice. The one legitimate path: Handshake's "school partnerships" feature, where some employers opt in to let students from *other* affiliated schools see their postings — worth checking in your own WSU Handshake account, but that's a manual thing only you can do logged into your own account, not something this scraper can automate.

### Product Manager / Business Analyst coverage
No new source was needed for this — it exposed a real bug instead. `HARD_AVOID_TITLE_RE`'s bare `manager` check (meant to catch seniority, e.g. "Engineering Manager") was killing every Product Manager posting outright, including ones explicitly labeled "Product Manager Summer 2027 Intern" — since "Manager" is also just the literal job title for PM roles at any level, not a seniority signal by itself. Fixed via the intern/co-op/fellowship/new-grad override described above. Also added `product management`/`product manager`/`roadmap`/`stakeholder` to `PROFILE["skills"]` (business-analyst/data-analytics terms were already there from an earlier round), and gave YC's Product category a fallback scoring text (`YC_CATEGORY_URLS` values) since the per-posting `roleType` field is frequently `null` in YC's own data.

## How filtering + scoring works

1. **Title filter** — only postings whose title contains `intern`, `co-op`, `new grad`, or `early career` survive at all (drops the senior/staff SWE postings that are also on these boards).
2. **Function filter** — drops recruiting/HR/sales/marketing postings that happen to say "early career" or "intern" but aren't engineering roles.
3. **Hard seniority / grad-level filter** — titles containing `senior`, `staff`, `principal`, `director`, `manager`, `lead`, `founding`, `head of`, `PhD`, `research scientist`, `master's student`, `doctoral`, `graduate researcher`, or `MBA` are dropped outright, no matter how many skills match — **unless** the title also explicitly says intern/co-op/fellowship/new-grad, which overrides it. That override matters: `manager` is also just the literal job title for Product roles at any level, so without it, "Product Manager Summer 2027 Intern" would get killed by the bare word "manager" despite plainly self-labeling as an internship (this was a real bug, caught and fixed). **Only undergraduate-eligible roles survive** — nothing requiring an in-progress or completed graduate degree.
4. **Domain exclusion** — cybersecurity (security engineer/analyst, infosec, pentesting, red/blue team), aerospace (aerospace, avionics, satellite, spacecraft, propulsion, flight software), and non-CS functions (finance, legal, accounting, medical/clinical, HR, sales, marketing) are dropped, checked against both title *and* description so a generic "Software Engineer Intern" on an aerospace team, or a "Medical Fellow" track sharing boilerplate text with an ML fellowship, still gets excluded.
5. **Skill requirement** — needs at least 2 real skill/keyword overlaps with your profile; a single generic hit (e.g. an accounting internship mentioning "Excel") isn't enough.
6. **US-only** — any posting whose location names a foreign country/city (London, Toronto, Bangalore, Sydney, etc.) is dropped outright, regardless of score.
7. **Scoring (0-100)** for everything that survives:
   - Skill/keyword overlap with your profile (including AI/data-analytics/business-analyst/ML terms) — 0-35 pts
   - Internship > fellowship > co-op > new-grad title strength — 0-25 pts
   - Seattle/WA or remote location — 0-20 pts
   - Recency (posted <7/14/30 days ago) — 0-20 pts
   - Penalties for "5+ years experience", "master's degree required", "security clearance", or any PhD mention — -15 pts each
8. Only postings scoring **55+** get written to the sheet at all — the point is a short, high-confidence list, not everything ranked. This bar is uniform across every source, including YC (see above for how that source compensates for structurally lower scores without lowering the bar itself).

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

`jobs.xlsx` has three tabs:
- **Internships** — postings whose title is literally an internship, co-op, or fellowship
- **Jobs** — everything else that qualified (new-grad titles, YC full-time roles) — same 55+ bar, just not internship-labeled
- **Startups** — a same-day mirror of both tabs above, narrowed to companies tagged `startup: True` **and** located in Seattle, the SF Bay Area, Texas, Arizona, or remote (`STARTUP_TAB_LOCATION_RE`). See "Startup-tag gap" above for how that tag gets set for companies outside the curated `COMPANIES` list.

Download `jobs.xlsx` from the repo and open in Excel/Google Sheets. Columns (same on all three tabs):

| Column | What it shows |
|---|---|
| Company Website | Click "Website →" — the company's homepage |
| Match % | Fit score, color coded |
| Fit | Strong Fit (80+) / Good Fit (65+) / Possible Fit (55+, floor of the sheet — re-tiered so all 3 labels stay reachable now that nothing below 55 is ever written) |
| Source | Which ATS it came from (gh / lever / ashby / simplify / yc) |
| Why It Matched | The specific skills/signals that drove the score |
| Link | Click "Apply →" to go straight to the posting |
| LinkedIn Search | Click "LinkedIn Search →" — opens a LinkedIn people search for `{Company} Washington State University`, logged in as you |
| LinkedIn Note Draft | A ready-to-send connection note (under 300 chars), tailored to the role — pick a person from the search above and paste this in |
| Email Draft | A longer version for email/InMail, with a subject line and your contact info already filled in |
| Status | Update yourself as you apply |
| Notes | Your own notes per job |

The current `jobs.xlsx` (every match found to date) is also attached directly to each Discord message, so you don't need to open GitHub to see the full list — just download the attachment from Discord.

### On the LinkedIn Search + message-draft columns — what this does and doesn't do
These are meant to speed up warm outreach, not automate it:
- **No LinkedIn scraping.** The link just opens LinkedIn's own people-search with useful filters pre-filled — you still browse the results and pick a real person yourself, logged into your own account. Automating LinkedIn scraping violates their ToS and risks your account getting flagged.
- **One combined link, not two.** This used to be two separate columns (a WSU-alumni search and a recruiter search); they're now one — a fellow WSU grad is a far better warm-intro target than a random recruiter, so the alumni search is "the better option" between them.
- **You still need to be logged into LinkedIn in your browser** for the search to show real results rather than a login wall — that's LinkedIn's requirement, not something this link can work around.
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
- `MIN_SCORE_TO_INCLUDE` — currently 55, uniform across every source (no per-source exceptions); raise it further if that's still surfacing too much, or lower it if the list feels too thin
- `ADDITIONAL_STARTUP_COMPANIES` — add a company here if a genuine startup discovered via SimplifyJobs isn't showing up on the Startups tab (see "Startup-tag gap" above)
- `NON_US_LOCATION_RE` — the sheet is US-only right now; remove a country/city from this pattern (or drop the check in `score_job` entirely) if you want to reopen it to a specific country
- `INTERNSHIP_TITLE_RE` — loosen this if you also want full-time junior/new-grad SWE roles, not just internships
