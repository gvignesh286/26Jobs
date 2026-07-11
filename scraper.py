"""
Giri Vignesh — Daily Internship Tracker
Pulls live postings directly from Greenhouse / Lever / Ashby job-board APIs
(no scraping, no blocking, no API key), filters to internship/co-op/
fellowship/new-grad roles, scores them against your resume profile, writes/
updates jobs.xlsx, and (if DISCORD_WEBHOOK_URL is set) posts a daily summary.
"""

import os
import re
import json
import time
from datetime import datetime, timezone
import requests
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Profile: Giri's skills and preferences ────────────────────────────────
PROFILE = {
    "skills": [
        "python", "javascript", "typescript", "react", "next.js", "nextjs",
        "flask", "node", "node.js", "express", "mongodb", "sql", "rest",
        "restful", "api", "html", "css", "docker", "git", "github", "aws",
        "tensorflow", "pytorch", "keras", "nlp", "llm", "rag",
        "retrieval augmented generation", "langchain", "embeddings",
        "vector database", "machine learning", "deep learning",
        "full stack", "full-stack", "backend", "frontend", "c++", "haskell",
        "agile", "ci/cd", "postman", "oop", "data structures", "algorithms",
        "web scraping", "beautifulsoup", "data analysis", "pandas",
        "data analyst", "data analytics", "business analyst",
        "data visualization", "tableau", "power bi", "excel", "kpi",
        "a/b testing", "statistics", "artificial intelligence",
    ],
    # Weighted higher than everything else — explicitly Seattle + remote per Giri's ask.
    "preferred_locations": [
        "seattle", "bellevue", "redmond", "kirkland", "washington",
        "pullman", "remote",
    ],
    # Soft penalties applied to description text — title-level seniority is
    # handled separately by HARD_AVOID_TITLE_RE (a hard drop, not a penalty).
    "avoid_signals": [
        "master's degree required", "msc required", "security clearance",
        "must be a us citizen", "citizenship required",
        "5+ years", "7+ years", "10+ years",
    ],
}

# Only these ATS platforms are queried — all are free, public, no-auth JSON
# APIs, chosen specifically to replace the unreliable Apify/LinkedIn scrape.
# "startup" tags which companies land on the separate "Startups" sheet tab —
# True for private/venture-funded companies, False for public/large-cap ones.
COMPANIES = [
    # ── Seattle-area ──
    {"name": "Smartsheet", "platform": "greenhouse", "token": "smartsheet", "startup": False},
    {"name": "Amperity", "platform": "greenhouse", "token": "amperity", "startup": True},
    {"name": "Textio", "platform": "greenhouse", "token": "textio", "startup": True},
    {"name": "Karat", "platform": "greenhouse", "token": "karat", "startup": True},
    {"name": "Xealth", "platform": "greenhouse", "token": "xealth", "startup": True},
    {"name": "Bungie", "platform": "greenhouse", "token": "bungie", "startup": False},
    {"name": "Rover", "platform": "lever", "token": "rover", "startup": False},
    {"name": "Outreach", "platform": "lever", "token": "outreach", "startup": True},
    {"name": "Highspot", "platform": "lever", "token": "highspot", "startup": True},
    {"name": "Qumulo", "platform": "ashby", "token": "qumulo", "startup": True},
    {"name": "PayScale", "platform": "ashby", "token": "payscale", "startup": True},
    {"name": "Adaptive Biotechnologies", "platform": "ashby", "token": "adaptive", "startup": False},

    # ── Remote-friendly / AI startups ──
    {"name": "Anthropic", "platform": "greenhouse", "token": "anthropic", "startup": True},
    {"name": "OpenAI", "platform": "ashby", "token": "openai", "startup": True},
    {"name": "Perplexity", "platform": "ashby", "token": "perplexity", "startup": True},
    {"name": "Harvey", "platform": "ashby", "token": "harvey", "startup": True},
    {"name": "Sierra", "platform": "ashby", "token": "sierra", "startup": True},
    {"name": "Decagon", "platform": "ashby", "token": "decagon", "startup": True},
    {"name": "Cursor (Anysphere)", "platform": "ashby", "token": "cursor", "startup": True},
    {"name": "Modal", "platform": "ashby", "token": "modal", "startup": True},
    {"name": "Baseten", "platform": "ashby", "token": "baseten", "startup": True},
    {"name": "Notion", "platform": "ashby", "token": "notion", "startup": True},
    {"name": "Ramp", "platform": "ashby", "token": "ramp", "startup": True},
    {"name": "Linear", "platform": "ashby", "token": "linear", "startup": True},
    {"name": "Replit", "platform": "ashby", "token": "replit", "startup": True},
    {"name": "PostHog", "platform": "ashby", "token": "posthog", "startup": True},
    {"name": "Runway", "platform": "ashby", "token": "runway", "startup": True},
    {"name": "Zapier", "platform": "ashby", "token": "zapier", "startup": True},
    {"name": "Airbyte", "platform": "ashby", "token": "airbyte", "startup": True},
    {"name": "Temporal", "platform": "ashby", "token": "temporal", "startup": True},
    {"name": "Substack", "platform": "ashby", "token": "substack", "startup": True},
    {"name": "Together AI", "platform": "greenhouse", "token": "togetherai", "startup": True},
    {"name": "Fireworks AI", "platform": "greenhouse", "token": "fireworksai", "startup": True},
    {"name": "Mercury", "platform": "greenhouse", "token": "mercury", "startup": True},
    {"name": "Vercel", "platform": "greenhouse", "token": "vercel", "startup": True},
    {"name": "Webflow", "platform": "greenhouse", "token": "webflow", "startup": True},
    {"name": "Airtable", "platform": "greenhouse", "token": "airtable", "startup": True},
    {"name": "GitLab", "platform": "greenhouse", "token": "gitlab", "startup": False},
    {"name": "Cockroach Labs", "platform": "greenhouse", "token": "cockroachlabs", "startup": True},
    {"name": "Scale AI", "platform": "greenhouse", "token": "scaleai", "startup": True},
    {"name": "Turing", "platform": "greenhouse", "token": "turing", "startup": True},
    {"name": "Flexport", "platform": "greenhouse", "token": "flexport", "startup": True},
    {"name": "Whoop", "platform": "lever", "token": "whoop", "startup": True},
    {"name": "Confluent", "platform": "ashby", "token": "confluent", "startup": False},
    {"name": "Plaid", "platform": "ashby", "token": "plaid", "startup": True},
    {"name": "Sift", "platform": "ashby", "token": "sift", "startup": True},

    # ── Larger tech (still worth a look — via their public ATS) ──
    {"name": "Stripe", "platform": "greenhouse", "token": "stripe", "startup": True},
    {"name": "Databricks", "platform": "greenhouse", "token": "databricks", "startup": True},
    {"name": "Figma", "platform": "greenhouse", "token": "figma", "startup": True},
    {"name": "Coinbase", "platform": "greenhouse", "token": "coinbase", "startup": False},
    {"name": "MongoDB", "platform": "greenhouse", "token": "mongodb", "startup": False},
    {"name": "Instacart", "platform": "greenhouse", "token": "instacart", "startup": False},
    {"name": "Reddit", "platform": "greenhouse", "token": "reddit", "startup": False},
    {"name": "Affirm", "platform": "greenhouse", "token": "affirm", "startup": False},
    {"name": "Asana", "platform": "greenhouse", "token": "asana", "startup": False},
    {"name": "Robinhood", "platform": "greenhouse", "token": "robinhood", "startup": False},
    {"name": "Discord", "platform": "greenhouse", "token": "discord", "startup": True},
    {"name": "Samsara", "platform": "greenhouse", "token": "samsara", "startup": False},
    {"name": "TripAdvisor", "platform": "greenhouse", "token": "tripadvisor", "startup": False},

    # ── Added to increase daily new-posting volume ──
    {"name": "Palantir", "platform": "lever", "token": "palantir", "startup": False},
    {"name": "Snowflake", "platform": "ashby", "token": "snowflake", "startup": False},
    {"name": "Cloudflare", "platform": "greenhouse", "token": "cloudflare", "startup": False},
    {"name": "Twilio", "platform": "greenhouse", "token": "twilio", "startup": False},
    {"name": "Duolingo", "platform": "greenhouse", "token": "duolingo", "startup": False},
    {"name": "Dropbox", "platform": "greenhouse", "token": "dropbox", "startup": False},
    {"name": "Pinterest", "platform": "greenhouse", "token": "pinterest", "startup": False},
    {"name": "Roblox", "platform": "greenhouse", "token": "roblox", "startup": False},
    {"name": "Okta", "platform": "greenhouse", "token": "okta", "startup": False},
    {"name": "Datadog", "platform": "greenhouse", "token": "datadog", "startup": False},
    {"name": "Elastic", "platform": "greenhouse", "token": "elastic", "startup": False},
    {"name": "PagerDuty", "platform": "greenhouse", "token": "pagerduty", "startup": False},
    {"name": "Amplitude", "platform": "greenhouse", "token": "amplitude", "startup": False},
    {"name": "Mixpanel", "platform": "greenhouse", "token": "mixpanel", "startup": True},
    {"name": "Braze", "platform": "greenhouse", "token": "braze", "startup": False},
    {"name": "Klaviyo", "platform": "greenhouse", "token": "klaviyo", "startup": False},
    {"name": "Miro", "platform": "ashby", "token": "miro", "startup": True},
    {"name": "Calendly", "platform": "greenhouse", "token": "calendly", "startup": True},
    {"name": "Chime", "platform": "greenhouse", "token": "chime", "startup": True},
    {"name": "SoFi", "platform": "greenhouse", "token": "sofi", "startup": False},
    {"name": "Wealthfront", "platform": "lever", "token": "wealthfront", "startup": True},
    {"name": "Betterment", "platform": "greenhouse", "token": "betterment", "startup": True},
    {"name": "Carta", "platform": "greenhouse", "token": "carta", "startup": True},
    {"name": "Vanta", "platform": "ashby", "token": "vanta", "startup": True},
    {"name": "Drata", "platform": "ashby", "token": "drata", "startup": True},
    {"name": "1Password", "platform": "ashby", "token": "1password", "startup": True},
    {"name": "LaunchDarkly", "platform": "greenhouse", "token": "launchdarkly", "startup": True},
    {"name": "Contentful", "platform": "greenhouse", "token": "contentful", "startup": True},
    {"name": "Sanity", "platform": "ashby", "token": "sanity", "startup": True},
    {"name": "Supabase", "platform": "ashby", "token": "supabase", "startup": True},
    {"name": "PlanetScale", "platform": "greenhouse", "token": "planetscale", "startup": True},
    {"name": "Neon", "platform": "lever", "token": "neon", "startup": True},
    {"name": "WorkOS", "platform": "ashby", "token": "workos", "startup": True},
    {"name": "Persona", "platform": "ashby", "token": "persona", "startup": True},
    {"name": "Merge", "platform": "ashby", "token": "merge", "startup": True},
    {"name": "Metronome", "platform": "greenhouse", "token": "metronome", "startup": True},
    {"name": "Column", "platform": "ashby", "token": "column", "startup": True},
    {"name": "Modern Treasury", "platform": "ashby", "token": "moderntreasury", "startup": True},
    {"name": "Lithic", "platform": "greenhouse", "token": "lithic", "startup": True},
    {"name": "Unit", "platform": "ashby", "token": "unit", "startup": True},
    {"name": "Alloy", "platform": "greenhouse", "token": "alloy", "startup": True},
    {"name": "Socure", "platform": "ashby", "token": "socure", "startup": True},
    {"name": "Fireblocks", "platform": "greenhouse", "token": "fireblocks", "startup": True},
    {"name": "Gemini", "platform": "greenhouse", "token": "gemini", "startup": True},
    {"name": "Anchorage", "platform": "lever", "token": "anchorage", "startup": True},
    {"name": "Wealthsimple", "platform": "ashby", "token": "wealthsimple", "startup": True},
    {"name": "Upstart", "platform": "greenhouse", "token": "upstart", "startup": False},
    {"name": "Block", "platform": "greenhouse", "token": "block", "startup": False},
    {"name": "Faire", "platform": "greenhouse", "token": "faire", "startup": True},
    {"name": "StockX", "platform": "greenhouse", "token": "stockx", "startup": True},
    {"name": "FanDuel", "platform": "greenhouse", "token": "fanduel", "startup": False},
    {"name": "PrizePicks", "platform": "greenhouse", "token": "prizepicks", "startup": True},
    {"name": "Sleeper", "platform": "ashby", "token": "sleeper", "startup": True},
    {"name": "Twitch", "platform": "greenhouse", "token": "twitch", "startup": False},
    {"name": "Cohere", "platform": "ashby", "token": "cohere", "startup": True},
    {"name": "LangChain", "platform": "ashby", "token": "langchain", "startup": True},
    {"name": "Pinecone", "platform": "ashby", "token": "pinecone", "startup": True},
    {"name": "Chroma", "platform": "ashby", "token": "trychroma", "startup": True},
    {"name": "IMC Trading", "platform": "greenhouse", "token": "imc", "startup": False},
    {"name": "Jane Street", "platform": "greenhouse", "token": "janestreet", "startup": False},
    {"name": "Squarespace", "platform": "greenhouse", "token": "squarespace", "startup": False},
    {"name": "Toast", "platform": "greenhouse", "token": "toast", "startup": False},
    {"name": "Postman", "platform": "greenhouse", "token": "postman", "startup": True},
    {"name": "Docker", "platform": "ashby", "token": "docker", "startup": True},
    {"name": "Render", "platform": "ashby", "token": "render", "startup": True},
    {"name": "Railway", "platform": "ashby", "token": "railway", "startup": True},
    {"name": "Warp", "platform": "greenhouse", "token": "warp", "startup": True},
]

# Only postings whose title looks like an internship / co-op / new-grad /
# fellowship role are kept — everything else on these boards (senior/staff
# SWE, etc.) is dropped before scoring so the sheet doesn't fill up with
# irrelevant rows. Fellowship programs (e.g. Anthropic Fellows Program,
# Scale AI's ML Fellow track) are treated the same as internships.
INTERNSHIP_TITLE_RE = re.compile(
    r"(?i)\b(intern(ship)?s?|co-?op|new grad|university grad|early career|"
    r"fellow(ship)?s?)\b"
)

# "Early career"/"intern"/"fellow" also gets used for recruiting, sales, HR,
# finance, and legal reqs on these boards (e.g. "Finance Fellow", "HRIS
# Analyst", "Legal Fellow") — exclude non-engineering/non-data functions
# outright so only AI/data-analytics/business-analyst/ML-flavored roles
# that are actually compatible with the resume survive.
NON_TECHNICAL_ROLE_RE = re.compile(
    r"(?i)\b(recruit(er|ing)?|people (strategy|ops|operations|analytics|technology)|"
    r"human resources|\bhr\b|hris|talent|human capital|sales|account executive|"
    r"marketing|content (writer|strategist)|copywriter|"
    r"finance|financial|accounting|accountant|payroll|legal|privacy|compliance|"
    r"audit|procurement|supply chain|customer (success|support)|workday|"
    r"medical|clinical)\b"
)

# Hard title exclusions — a rising-junior CS undergrad isn't eligible for
# these regardless of keyword/skill overlap, so drop them before scoring.
# Also covers grad-degree-track titles (Giri is undergrad-only).
HARD_AVOID_TITLE_RE = re.compile(
    r"(?i)\b(senior|staff|principal|director|manager|"
    r"phd|ph\.d\.|postdoc(toral)?|research scientist|"
    r"graduate program|master'?s?\s+student|phd student|"
    r"doctoral( student)?|graduate researcher|mba)\b"
)

# Whole-domain exclusions — checked against title+description since these
# fields sometimes only show up in the description, not the title. Not a
# fit regardless of skill overlap.
DOMAIN_EXCLUDE_RE = re.compile(
    r"(?i)\b(cybersecurity|cyber security|security engineer|security analyst|"
    r"infosec|information security|penetration test(ing|er)?|red team|"
    r"blue team|soc analyst|application security|"
    r"aerospace|avionics|satellite|spacecraft|propulsion|flight software|aircraft)\b"
)

# US-only — checked against the location field. Blocklist approach (rather
# than a US-city whitelist) since ATS location strings are free-text and a
# whitelist would miss valid US cities; every non-US posting seen in testing
# named a specific foreign country/city, so blocking those is reliable.
NON_US_LOCATION_RE = re.compile(
    r"(?i)\b(united kingdom|\buk\b|london|belgrade|serbia|sydney|melbourne|"
    r"australia|canada|toronto|vancouver|montreal|india|bangalore|delhi|"
    r"mumbai|hyderabad|singapore|germany|berlin|munich|france|paris|"
    r"ireland|dublin|netherlands|amsterdam|spain|madrid|barcelona|italy|"
    r"milan|rome|japan|tokyo|china|beijing|shanghai|brazil|sao paulo|"
    r"mexico|poland|warsaw|portugal|lisbon|switzerland|zurich|sweden|"
    r"stockholm|denmark|copenhagen|norway|oslo|israel|tel aviv|"
    r"philippines|manila|vietnam|indonesia|jakarta|malaysia|"
    r"south africa|new zealand|austria|vienna|belgium|brussels)\b"
)

MIN_SCORE_TO_INCLUDE = 50  # below this, a posting is dropped entirely (not just low-ranked)


def skill_pattern(term: str) -> re.Pattern:
    """Word-boundary-safe match so short/symbol-heavy skills (e.g. 'rag', 'api',
    'oop') don't false-positive inside unrelated words ('storage', 'capital',
    'cooperate')."""
    return re.compile(rf"(?<![a-z0-9]){re.escape(term)}(?![a-z0-9])")


SKILL_PATTERNS = [(s, skill_pattern(s)) for s in PROFILE["skills"]]
AVOID_PATTERNS = [(s, skill_pattern(s)) for s in PROFILE["avoid_signals"]]


# ── Fetchers ──────────────────────────────────────────────────────────────
def strip_html(html: str) -> str:
    text = re.sub(r"<[^>]+>", " ", html or "")
    text = text.replace("&nbsp;", " ").replace("&amp;", "&")
    return re.sub(r"\s+", " ", text).strip()


def fetch_greenhouse(company: dict) -> list[dict]:
    token = company["token"]
    url = f"https://boards-api.greenhouse.io/v1/boards/{token}/jobs?content=true"
    r = requests.get(url, timeout=15)
    r.raise_for_status()
    jobs = []
    for j in r.json().get("jobs", []):
        offices = j.get("offices") or []
        location = j.get("location", {}).get("name") or (offices[0]["name"] if offices else "")
        jobs.append({
            "id": f"gh:{token}:{j['id']}",
            "title": j.get("title", ""),
            "companyName": company["name"],
            "isStartup": company["startup"],
            "location": location,
            "descriptionText": strip_html(j.get("content", "")),
            "link": j.get("absolute_url", ""),
            "postedAt": j.get("updated_at", ""),
        })
    return jobs


def fetch_lever(company: dict) -> list[dict]:
    token = company["token"]
    url = f"https://api.lever.co/v0/postings/{token}?mode=json"
    r = requests.get(url, timeout=15)
    r.raise_for_status()
    jobs = []
    for j in r.json():
        cats = j.get("categories", {}) or {}
        jobs.append({
            "id": f"lever:{token}:{j['id']}",
            "title": j.get("text", ""),
            "companyName": company["name"],
            "isStartup": company["startup"],
            "location": cats.get("location", ""),
            "descriptionText": strip_html(j.get("descriptionPlain") or j.get("description", "")),
            "link": j.get("hostedUrl", ""),
            "postedAt": datetime.fromtimestamp(
                j.get("createdAt", 0) / 1000, tz=timezone.utc
            ).isoformat() if j.get("createdAt") else "",
        })
    return jobs


def fetch_ashby(company: dict) -> list[dict]:
    token = company["token"]
    url = f"https://api.ashbyhq.com/posting-api/job-board/{token}"
    r = requests.get(url, timeout=15)
    r.raise_for_status()
    jobs = []
    for j in r.json().get("jobs", []):
        location = j.get("location", "")
        if j.get("isRemote") and "remote" not in location.lower():
            location = f"{location} (Remote)".strip()
        jobs.append({
            "id": f"ashby:{token}:{j['id']}",
            "title": j.get("title", ""),
            "companyName": company["name"],
            "isStartup": company["startup"],
            "location": location,
            "descriptionText": strip_html(j.get("descriptionHtml") or ""),
            "link": j.get("jobUrl", ""),
            "postedAt": j.get("publishedAt", ""),
        })
    return jobs


FETCHERS = {"greenhouse": fetch_greenhouse, "lever": fetch_lever, "ashby": fetch_ashby}


def scrape_jobs() -> list[dict]:
    all_jobs = []
    for company in COMPANIES:
        fetcher = FETCHERS[company["platform"]]
        try:
            jobs = fetcher(company)
            internship_jobs = [
                j for j in jobs
                if INTERNSHIP_TITLE_RE.search(j["title"])
                and not NON_TECHNICAL_ROLE_RE.search(j["title"])
            ]
            print(f"{company['name']:28s} ({company['platform']:10s}): "
                  f"{len(jobs):4d} total, {len(internship_jobs):3d} internship-tagged")
            all_jobs.extend(internship_jobs)
        except Exception as e:
            print(f"{company['name']:28s} ({company['platform']:10s}): FAILED — {e}")
        time.sleep(0.3)  # be polite to free public APIs
    return all_jobs


# ── Scoring ────────────────────────────────────────────────────────────────
def score_job(job: dict) -> tuple[int, list[str]]:
    """
    Returns (score 0-100, list of matched reasons).
    Breakdown:
      - Skill keyword match   : 0-35 pts
      - Title / role fit      : 0-25 pts
      - Location preference   : 0-20 pts
      - Recency                : 0-20 pts
      - Negative signals       : subtracted, floor 0
    """
    title = (job.get("title") or "").lower()
    location = (job.get("location") or "").lower()
    description = (job.get("descriptionText") or "").lower()
    full_text = f"{title} {description}"

    # ── Hard title filter — drop regardless of skill/keyword overlap ──────
    if HARD_AVOID_TITLE_RE.search(title):
        return 0, [f"filtered: '{HARD_AVOID_TITLE_RE.search(title).group(0)}' in title"]

    # ── Domain exclusion — cybersecurity / aerospace, title or description ─
    domain_hit = DOMAIN_EXCLUDE_RE.search(full_text)
    if domain_hit:
        return 0, [f"filtered: '{domain_hit.group(0)}' (excluded domain)"]

    # ── US-only — drop anything whose location names a foreign country/city ─
    non_us_hit = NON_US_LOCATION_RE.search(location)
    if non_us_hit:
        return 0, [f"filtered: '{non_us_hit.group(0)}' (non-US location)"]

    reasons = []
    score = 0

    # ── Skill match (0-35) ────────────────────────────────────────────────
    matched_skills = [s for s, pat in SKILL_PATTERNS if pat.search(full_text)]
    if len(matched_skills) < 2:
        # A single keyword hit is too weak a signal on its own (e.g. an
        # accounting internship mentioning "Excel", or boilerplate template
        # text a company reuses across unrelated fellowship tracks) —
        # require at least 2 real overlaps before treating it as a fit.
        return 0, ["filtered: insufficient technical skill overlap"]
    skill_score = min(35, len(matched_skills) * 3)
    score += skill_score
    reasons.append(f"Skills: {', '.join(matched_skills[:6])}")

    # ── Title / role fit (0-25) ───────────────────────────────────────────
    if re.search(r"\bintern(ship)?s?\b", title):
        score += 25
        reasons.append("Internship title")
    elif re.search(r"fellow(ship)?s?", title):
        score += 23
        reasons.append("Fellowship title")
    elif re.search(r"co-?op", title):
        score += 20
        reasons.append("Co-op title")
    elif re.search(r"new grad|university grad|early career", title):
        score += 12
        reasons.append("New-grad title")

    # ── Location (0-20) ───────────────────────────────────────────────────
    for loc in PROFILE["preferred_locations"]:
        if loc in location:
            score += 20
            reasons.append(f"Location: {job.get('location')}")
            break

    # ── Recency (0-20) ────────────────────────────────────────────────────
    posted_at = job.get("postedAt")
    if posted_at:
        try:
            posted_dt = datetime.fromisoformat(posted_at.replace("Z", "+00:00"))
            days_old = (datetime.now(timezone.utc) - posted_dt).days
            if days_old <= 7:
                score += 20
                reasons.append("Posted <7d ago")
            elif days_old <= 14:
                score += 14
            elif days_old <= 30:
                score += 8
            else:
                score += 2
        except Exception:
            pass

    # ── Negative signals ──────────────────────────────────────────────────
    penalties = [bad for bad, pat in AVOID_PATTERNS if pat.search(full_text)]
    if re.search(r"(?<![a-z])phd(?![a-z])|ph\.d\.", full_text):
        penalties.append("PhD-oriented role")
    for _ in penalties:
        score -= 15
    if penalties:
        reasons.append(f"Flags: {', '.join(penalties[:3])}")

    return max(0, min(score, 100)), reasons


def chance_label(score: int) -> str:
    if score >= 70: return "Strong Fit"
    if score >= 55: return "Good Fit"
    return "Possible Fit"


def chance_color(score: int) -> str:
    if score >= 70: return "C6EFCE"   # green
    if score >= 55: return "FFEB9C"   # yellow
    return "FFCC99"                    # orange


# ── Deduplicate against existing sheet ───────────────────────────────────
MAIN_SHEET_NAME = "Internship Matches"
STARTUP_SHEET_NAME = "Startups"


def load_existing_ids(path: str) -> set:
    if not os.path.exists(path):
        return set()
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb[MAIN_SHEET_NAME] if MAIN_SHEET_NAME in wb.sheetnames else wb.active
        return {str(row[0]) for row in ws.iter_rows(min_row=2, values_only=True) if row[0]}
    except Exception:
        return set()


# ── Excel builder ─────────────────────────────────────────────────────────
HEADERS = [
    "Job ID", "Date Found", "Title", "Company", "Location",
    "Match %", "Fit", "Source", "Why It Matched", "Link", "Status", "Notes"
]

COL_WIDTHS = [24, 12, 34, 24, 20, 10, 13, 10, 40, 14, 14, 24]

NAVY = "1A3A5C"
WHITE = "FFFFFF"
LIGHT = "F2F5F9"
BORDER_COLOR = "CCCCCC"


def thin_border():
    s = Side(style="thin", color=BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)


def write_excel(jobs_scored: list[dict], path: str) -> list[tuple[dict, int]]:
    """Returns the (job, score) pairs actually written to the main sheet —
    used to build the Discord notification so it only reports what's
    genuinely new today. The Startups sheet is a same-day mirror subset."""
    existing_ids = load_existing_ids(path)
    added = []

    if os.path.exists(path):
        wb = openpyxl.load_workbook(path)
        ws_main = wb[MAIN_SHEET_NAME] if MAIN_SHEET_NAME in wb.sheetnames else wb.active
        ws_main.title = MAIN_SHEET_NAME
        ws_startup = wb[STARTUP_SHEET_NAME] if STARTUP_SHEET_NAME in wb.sheetnames else None
        if ws_startup is None:
            ws_startup = wb.create_sheet(STARTUP_SHEET_NAME)
            _write_header(ws_startup)
    else:
        wb = openpyxl.Workbook()
        ws_main = wb.active
        ws_main.title = MAIN_SHEET_NAME
        _write_header(ws_main)
        ws_startup = wb.create_sheet(STARTUP_SHEET_NAME)
        _write_header(ws_startup)

    for job in jobs_scored:
        if job["id"] in existing_ids:
            continue
        score = _append_row(ws_main, job)
        if score is not None:
            added.append((job, score))
            if job.get("isStartup"):
                _append_row(ws_startup, job)

    print(f"Added {len(added)} new internship postings "
          f"({sum(1 for j, _ in added if j.get('isStartup'))} startup) to the sheet.")

    for ws in (ws_main, ws_startup):
        ws.freeze_panes = "A2"
        for i, width in enumerate(COL_WIDTHS, 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        ws.auto_filter.ref = ws.dimensions

    wb.save(path)
    print(f"Saved to {path}")
    return added


def _write_header(ws):
    for col, header in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color=WHITE, name="Arial", size=10)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border()
    ws.row_dimensions[1].height = 22


def _append_row(ws, job: dict) -> int | None:
    score, reasons = score_job(job)
    if score < MIN_SCORE_TO_INCLUDE:
        return None  # not a strong enough fit — don't clutter the sheet

    row = ws.max_row + 1
    fill_color = LIGHT if row % 2 == 0 else WHITE
    source = job["id"].split(":")[0].capitalize()

    values = [
        job["id"],
        datetime.today().strftime("%Y-%m-%d"),
        job.get("title", ""),
        job.get("companyName", ""),
        job.get("location", ""),
        score,
        chance_label(score),
        source,
        ", ".join(reasons[:4]),
        job.get("link", ""),
        "Not Applied",
        "",
    ]

    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = Font(name="Arial", size=9)
        cell.alignment = Alignment(vertical="center", wrap_text=(col in [3, 9]))
        cell.border = thin_border()
        if col not in [6, 7]:
            cell.fill = PatternFill("solid", fgColor=fill_color)

    pct_cell = ws.cell(row=row, column=6)
    pct_cell.value = f"{score}%"
    pct_cell.font = Font(name="Arial", size=9, bold=True)
    pct_cell.fill = PatternFill("solid", fgColor=chance_color(score))
    pct_cell.alignment = Alignment(horizontal="center", vertical="center")

    ch_cell = ws.cell(row=row, column=7)
    ch_cell.fill = PatternFill("solid", fgColor=chance_color(score))
    ch_cell.alignment = Alignment(horizontal="center", vertical="center")

    link = values[9]
    if link:
        link_cell = ws.cell(row=row, column=10, value="Apply →")
        link_cell.hyperlink = link
        link_cell.font = Font(name="Arial", size=9, color="185FA5", underline="single")
        link_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[row].height = 18
    return score


# ── Discord notification ──────────────────────────────────────────────────
def notify_discord(webhook_url: str, added: list[tuple[dict, int]], companies_scanned: int, xlsx_path: str):
    today = datetime.today().strftime("%Y-%m-%d")

    if not added:
        payload = {
            "embeds": [{
                "title": "🎯 Daily Internship Tracker",
                "description": f"Ran today ({today}) — no new matches across {companies_scanned} companies. Full sheet attached below.",
                "color": 0x1A3A5C,
            }]
        }
    else:
        added_sorted = sorted(added, key=lambda pair: pair[1], reverse=True)
        lines = []
        for job, score in added_sorted[:10]:
            fit = chance_label(score)
            lines.append(
                f"**{score}% {fit}** — {job.get('title', '')} @ {job.get('companyName', '')}\n"
                f"[Apply]({job.get('link', '')})"
            )
        description = "\n\n".join(lines)
        if len(added_sorted) > 10:
            description += f"\n\n...and {len(added_sorted) - 10} more in the sheet."
        description += "\n\nFull sheet (all matches to date) attached below."

        payload = {
            "embeds": [{
                "title": f"🎯 {len(added)} New Internship/Fellowship Match(es) — {today}",
                "description": description,
                "color": 0x2ECC71,
                "footer": {"text": f"Scanned {companies_scanned} companies"},
            }]
        }

    try:
        if os.path.exists(xlsx_path):
            with open(xlsx_path, "rb") as f:
                r = requests.post(
                    webhook_url,
                    data={"payload_json": json.dumps(payload)},
                    files={"file": ("jobs.xlsx", f,
                                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                    timeout=30,
                )
        else:
            r = requests.post(webhook_url, json=payload, timeout=15)
        r.raise_for_status()
        print("Discord notification sent (with jobs.xlsx attached).")
    except Exception as e:
        print(f"Discord notification FAILED (non-fatal): {e}")


# ── Main ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    output_path = os.path.join(os.path.dirname(__file__), "jobs.xlsx")

    jobs = scrape_jobs()
    print(f"\n{len(jobs)} internship-tagged postings found across {len(COMPANIES)} companies.")

    added = write_excel(jobs, output_path)

    webhook_url = os.environ.get("DISCORD_WEBHOOK_URL")
    if webhook_url:
        notify_discord(webhook_url, added, len(COMPANIES), output_path)
    else:
        print("DISCORD_WEBHOOK_URL not set — skipping Discord notification.")

    print(f"\nDone — {datetime.today().strftime('%Y-%m-%d %H:%M')} PST")
