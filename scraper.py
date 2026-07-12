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
import html
from datetime import datetime, timezone
from urllib.parse import quote
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
        "product management", "product manager", "roadmap", "stakeholder",
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

CONTACT = {
    "name": "Giri Vignesh",
    "school": "Washington State University",
    "grad": "May 2027",
    "email": "girivignesh5@gmail.com",
    "phone": "(425) 362-2938",
    "github": "github.com/gvignesh286",
}

# One-line project highlights, keyed by which skill category they best sell.
# _pick_highlight() matches these against a job's matched skills so the
# outreach draft references the most relevant project, not a generic one.
HIGHLIGHTS = {
    "ml": {
        "keywords": ["tensorflow", "pytorch", "keras", "nlp", "llm", "rag",
                      "langchain", "embeddings", "machine learning", "deep learning",
                      "artificial intelligence"],
        "short": "I recently built APEX, an ML-driven trading signal system with a self-retraining model",
        "email": "Most recently I built APEX, a full-stack system that combines a technical-analysis "
                  "signal engine with a RandomForest model that retrains itself on live trading outcomes "
                  "— deployed end-to-end on AWS.",
    },
    "web": {
        "keywords": ["react", "next.js", "nextjs", "flask", "node", "node.js", "express",
                      "mongodb", "full stack", "full-stack", "backend", "frontend", "rest", "restful"],
        "short": "I've shipped a few full-stack apps (Flask/React, deployed on AWS) including an AI travel planner",
        "email": "I've shipped several full-stack projects — including TravelBuddy, an AI-powered travel "
                  "planner with a Flask backend and React frontend deployed on AWS, and a hackathon project "
                  "integrating the Anthropic Claude API built in 24 hours.",
    },
    "data": {
        "keywords": ["data analysis", "data analyst", "data analytics", "business analyst",
                      "pandas", "data visualization", "statistics", "sql"],
        "short": "I've built a couple of data-heavy projects, including an automated stock signal analyzer",
        "email": "I've built a couple of data-focused projects — an automated stock analysis tool applying "
                  "technical indicators to live market data, and an interactive trail-mapping app processing "
                  "200+ real-world datasets with Pandas.",
    },
    "default": {
        "keywords": [],
        "short": "I'm a CS student who's shipped several full-stack and ML projects, including one at a hackathon",
        "email": "I've spent the last couple years building full-stack and ML projects — from an AI travel "
                  "planner to a hackathon project (top result at WSU's CrimsonCode, 250+ participants) "
                  "integrating the Anthropic Claude API.",
    },
}

# Only these ATS platforms are queried — all are free, public, no-auth JSON
# APIs, chosen specifically to replace the unreliable Apify/LinkedIn scrape.
# "startup" tags which companies land on the separate "Startups" sheet tab —
# True for private/venture-funded companies, False for public/large-cap ones.
COMPANIES = [
    # ── Seattle-area ──
    {"name": "Smartsheet", "platform": "greenhouse", "token": "smartsheet", "startup": False, "website": "https://www.smartsheet.com"},
    {"name": "Amperity", "platform": "greenhouse", "token": "amperity", "startup": True, "website": "https://amperity.com"},
    {"name": "Textio", "platform": "greenhouse", "token": "textio", "startup": True, "website": "https://textio.com"},
    {"name": "Karat", "platform": "greenhouse", "token": "karat", "startup": True, "website": "https://karat.com"},
    {"name": "Xealth", "platform": "greenhouse", "token": "xealth", "startup": True, "website": "https://xealth.com"},
    {"name": "Bungie", "platform": "greenhouse", "token": "bungie", "startup": False, "website": "https://www.bungie.net"},
    {"name": "Rover", "platform": "lever", "token": "rover", "startup": False, "website": "https://www.rover.com"},
    {"name": "Outreach", "platform": "lever", "token": "outreach", "startup": True, "website": "https://www.outreach.io"},
    {"name": "Highspot", "platform": "lever", "token": "highspot", "startup": True, "website": "https://www.highspot.com"},
    {"name": "Qumulo", "platform": "ashby", "token": "qumulo", "startup": True, "website": "https://qumulo.com"},
    {"name": "PayScale", "platform": "ashby", "token": "payscale", "startup": True, "website": "https://www.payscale.com"},
    {"name": "Adaptive Biotechnologies", "platform": "ashby", "token": "adaptive", "startup": False, "website": "https://www.adaptivebiotech.com"},

    # ── Remote-friendly / AI startups ──
    {"name": "Anthropic", "platform": "greenhouse", "token": "anthropic", "startup": True, "website": "https://www.anthropic.com"},
    {"name": "OpenAI", "platform": "ashby", "token": "openai", "startup": True, "website": "https://openai.com"},
    {"name": "Perplexity", "platform": "ashby", "token": "perplexity", "startup": True, "website": "https://www.perplexity.ai"},
    {"name": "Harvey", "platform": "ashby", "token": "harvey", "startup": True, "website": "https://www.harvey.ai"},
    {"name": "Sierra", "platform": "ashby", "token": "sierra", "startup": True, "website": "https://sierra.ai"},
    {"name": "Decagon", "platform": "ashby", "token": "decagon", "startup": True, "website": "https://decagon.ai"},
    {"name": "Cursor (Anysphere)", "platform": "ashby", "token": "cursor", "startup": True, "website": "https://www.cursor.com"},
    {"name": "Modal", "platform": "ashby", "token": "modal", "startup": True, "website": "https://modal.com"},
    {"name": "Baseten", "platform": "ashby", "token": "baseten", "startup": True, "website": "https://www.baseten.co"},
    {"name": "Notion", "platform": "ashby", "token": "notion", "startup": True, "website": "https://www.notion.so"},
    {"name": "Ramp", "platform": "ashby", "token": "ramp", "startup": True, "website": "https://ramp.com"},
    {"name": "Linear", "platform": "ashby", "token": "linear", "startup": True, "website": "https://linear.app"},
    {"name": "Replit", "platform": "ashby", "token": "replit", "startup": True, "website": "https://replit.com"},
    {"name": "PostHog", "platform": "ashby", "token": "posthog", "startup": True, "website": "https://posthog.com"},
    {"name": "Runway", "platform": "ashby", "token": "runway", "startup": True, "website": "https://runwayml.com"},
    {"name": "Zapier", "platform": "ashby", "token": "zapier", "startup": True, "website": "https://zapier.com"},
    {"name": "Airbyte", "platform": "ashby", "token": "airbyte", "startup": True, "website": "https://airbyte.com"},
    {"name": "Temporal", "platform": "ashby", "token": "temporal", "startup": True, "website": "https://temporal.io"},
    {"name": "Substack", "platform": "ashby", "token": "substack", "startup": True, "website": "https://substack.com"},
    {"name": "Together AI", "platform": "greenhouse", "token": "togetherai", "startup": True, "website": "https://www.together.ai"},
    {"name": "Fireworks AI", "platform": "greenhouse", "token": "fireworksai", "startup": True, "website": "https://fireworks.ai"},
    {"name": "Mercury", "platform": "greenhouse", "token": "mercury", "startup": True, "website": "https://mercury.com"},
    {"name": "Vercel", "platform": "greenhouse", "token": "vercel", "startup": True, "website": "https://vercel.com"},
    {"name": "Webflow", "platform": "greenhouse", "token": "webflow", "startup": True, "website": "https://webflow.com"},
    {"name": "Airtable", "platform": "greenhouse", "token": "airtable", "startup": True, "website": "https://www.airtable.com"},
    {"name": "GitLab", "platform": "greenhouse", "token": "gitlab", "startup": False, "website": "https://about.gitlab.com"},
    {"name": "Cockroach Labs", "platform": "greenhouse", "token": "cockroachlabs", "startup": True, "website": "https://www.cockroachlabs.com"},
    {"name": "Scale AI", "platform": "greenhouse", "token": "scaleai", "startup": True, "website": "https://scale.com"},
    {"name": "Turing", "platform": "greenhouse", "token": "turing", "startup": True, "website": "https://www.turing.com"},
    {"name": "Flexport", "platform": "greenhouse", "token": "flexport", "startup": True, "website": "https://www.flexport.com"},
    {"name": "Whoop", "platform": "lever", "token": "whoop", "startup": True, "website": "https://www.whoop.com"},
    {"name": "Confluent", "platform": "ashby", "token": "confluent", "startup": False, "website": "https://www.confluent.io"},
    {"name": "Plaid", "platform": "ashby", "token": "plaid", "startup": True, "website": "https://plaid.com"},
    {"name": "Sift", "platform": "ashby", "token": "sift", "startup": True, "website": "https://sift.com"},

    # ── Larger tech (still worth a look — via their public ATS) ──
    {"name": "Stripe", "platform": "greenhouse", "token": "stripe", "startup": True, "website": "https://stripe.com"},
    {"name": "Databricks", "platform": "greenhouse", "token": "databricks", "startup": True, "website": "https://www.databricks.com"},
    {"name": "Figma", "platform": "greenhouse", "token": "figma", "startup": True, "website": "https://www.figma.com"},
    {"name": "Coinbase", "platform": "greenhouse", "token": "coinbase", "startup": False, "website": "https://www.coinbase.com"},
    {"name": "MongoDB", "platform": "greenhouse", "token": "mongodb", "startup": False, "website": "https://www.mongodb.com"},
    {"name": "Instacart", "platform": "greenhouse", "token": "instacart", "startup": False, "website": "https://www.instacart.com"},
    {"name": "Reddit", "platform": "greenhouse", "token": "reddit", "startup": False, "website": "https://www.redditinc.com"},
    {"name": "Affirm", "platform": "greenhouse", "token": "affirm", "startup": False, "website": "https://www.affirm.com"},
    {"name": "Asana", "platform": "greenhouse", "token": "asana", "startup": False, "website": "https://asana.com"},
    {"name": "Robinhood", "platform": "greenhouse", "token": "robinhood", "startup": False, "website": "https://robinhood.com"},
    {"name": "Discord", "platform": "greenhouse", "token": "discord", "startup": True, "website": "https://discord.com"},
    {"name": "Samsara", "platform": "greenhouse", "token": "samsara", "startup": False, "website": "https://www.samsara.com"},
    {"name": "TripAdvisor", "platform": "greenhouse", "token": "tripadvisor", "startup": False, "website": "https://www.tripadvisor.com"},

    # ── Added to increase daily new-posting volume ──
    {"name": "Palantir", "platform": "lever", "token": "palantir", "startup": False, "website": "https://www.palantir.com"},
    {"name": "Snowflake", "platform": "ashby", "token": "snowflake", "startup": False, "website": "https://www.snowflake.com"},
    {"name": "Cloudflare", "platform": "greenhouse", "token": "cloudflare", "startup": False, "website": "https://www.cloudflare.com"},
    {"name": "Twilio", "platform": "greenhouse", "token": "twilio", "startup": False, "website": "https://www.twilio.com"},
    {"name": "Duolingo", "platform": "greenhouse", "token": "duolingo", "startup": False, "website": "https://www.duolingo.com"},
    {"name": "Dropbox", "platform": "greenhouse", "token": "dropbox", "startup": False, "website": "https://www.dropbox.com"},
    {"name": "Pinterest", "platform": "greenhouse", "token": "pinterest", "startup": False, "website": "https://www.pinterest.com"},
    {"name": "Roblox", "platform": "greenhouse", "token": "roblox", "startup": False, "website": "https://www.roblox.com"},
    {"name": "Okta", "platform": "greenhouse", "token": "okta", "startup": False, "website": "https://www.okta.com"},
    {"name": "Datadog", "platform": "greenhouse", "token": "datadog", "startup": False, "website": "https://www.datadoghq.com"},
    {"name": "Elastic", "platform": "greenhouse", "token": "elastic", "startup": False, "website": "https://www.elastic.co"},
    {"name": "PagerDuty", "platform": "greenhouse", "token": "pagerduty", "startup": False, "website": "https://www.pagerduty.com"},
    {"name": "Amplitude", "platform": "greenhouse", "token": "amplitude", "startup": False, "website": "https://amplitude.com"},
    {"name": "Mixpanel", "platform": "greenhouse", "token": "mixpanel", "startup": True, "website": "https://mixpanel.com"},
    {"name": "Braze", "platform": "greenhouse", "token": "braze", "startup": False, "website": "https://www.braze.com"},
    {"name": "Klaviyo", "platform": "greenhouse", "token": "klaviyo", "startup": False, "website": "https://www.klaviyo.com"},
    {"name": "Miro", "platform": "ashby", "token": "miro", "startup": True, "website": "https://miro.com"},
    {"name": "Calendly", "platform": "greenhouse", "token": "calendly", "startup": True, "website": "https://calendly.com"},
    {"name": "Chime", "platform": "greenhouse", "token": "chime", "startup": True, "website": "https://www.chime.com"},
    {"name": "SoFi", "platform": "greenhouse", "token": "sofi", "startup": False, "website": "https://www.sofi.com"},
    {"name": "Wealthfront", "platform": "lever", "token": "wealthfront", "startup": True, "website": "https://www.wealthfront.com"},
    {"name": "Betterment", "platform": "greenhouse", "token": "betterment", "startup": True, "website": "https://www.betterment.com"},
    {"name": "Carta", "platform": "greenhouse", "token": "carta", "startup": True, "website": "https://carta.com"},
    {"name": "Vanta", "platform": "ashby", "token": "vanta", "startup": True, "website": "https://www.vanta.com"},
    {"name": "Drata", "platform": "ashby", "token": "drata", "startup": True, "website": "https://drata.com"},
    {"name": "1Password", "platform": "ashby", "token": "1password", "startup": True, "website": "https://1password.com"},
    {"name": "LaunchDarkly", "platform": "greenhouse", "token": "launchdarkly", "startup": True, "website": "https://launchdarkly.com"},
    {"name": "Contentful", "platform": "greenhouse", "token": "contentful", "startup": True, "website": "https://www.contentful.com"},
    {"name": "Sanity", "platform": "ashby", "token": "sanity", "startup": True, "website": "https://www.sanity.io"},
    {"name": "Supabase", "platform": "ashby", "token": "supabase", "startup": True, "website": "https://supabase.com"},
    {"name": "PlanetScale", "platform": "greenhouse", "token": "planetscale", "startup": True, "website": "https://planetscale.com"},
    {"name": "Neon", "platform": "lever", "token": "neon", "startup": True, "website": "https://neon.tech"},
    {"name": "WorkOS", "platform": "ashby", "token": "workos", "startup": True, "website": "https://workos.com"},
    {"name": "Persona", "platform": "ashby", "token": "persona", "startup": True, "website": "https://withpersona.com"},
    {"name": "Merge", "platform": "ashby", "token": "merge", "startup": True, "website": "https://www.merge.dev"},
    {"name": "Metronome", "platform": "greenhouse", "token": "metronome", "startup": True, "website": "https://metronome.com"},
    {"name": "Column", "platform": "ashby", "token": "column", "startup": True, "website": "https://column.com"},
    {"name": "Modern Treasury", "platform": "ashby", "token": "moderntreasury", "startup": True, "website": "https://www.moderntreasury.com"},
    {"name": "Lithic", "platform": "greenhouse", "token": "lithic", "startup": True, "website": "https://www.lithic.com"},
    {"name": "Unit", "platform": "ashby", "token": "unit", "startup": True, "website": "https://www.unit.co"},
    {"name": "Alloy", "platform": "greenhouse", "token": "alloy", "startup": True, "website": "https://www.alloy.com"},
    {"name": "Socure", "platform": "ashby", "token": "socure", "startup": True, "website": "https://www.socure.com"},
    {"name": "Fireblocks", "platform": "greenhouse", "token": "fireblocks", "startup": True, "website": "https://www.fireblocks.com"},
    {"name": "Gemini", "platform": "greenhouse", "token": "gemini", "startup": True, "website": "https://www.gemini.com"},
    {"name": "Anchorage", "platform": "lever", "token": "anchorage", "startup": True, "website": "https://www.anchorage.com"},
    {"name": "Wealthsimple", "platform": "ashby", "token": "wealthsimple", "startup": True, "website": "https://www.wealthsimple.com"},
    {"name": "Upstart", "platform": "greenhouse", "token": "upstart", "startup": False, "website": "https://www.upstart.com"},
    {"name": "Block", "platform": "greenhouse", "token": "block", "startup": False, "website": "https://block.xyz"},
    {"name": "Faire", "platform": "greenhouse", "token": "faire", "startup": True, "website": "https://www.faire.com"},
    {"name": "StockX", "platform": "greenhouse", "token": "stockx", "startup": True, "website": "https://stockx.com"},
    {"name": "FanDuel", "platform": "greenhouse", "token": "fanduel", "startup": False, "website": "https://www.fanduel.com"},
    {"name": "PrizePicks", "platform": "greenhouse", "token": "prizepicks", "startup": True, "website": "https://www.prizepicks.com"},
    {"name": "Sleeper", "platform": "ashby", "token": "sleeper", "startup": True, "website": "https://sleeper.com"},
    {"name": "Twitch", "platform": "greenhouse", "token": "twitch", "startup": False, "website": "https://www.twitch.tv"},
    {"name": "Cohere", "platform": "ashby", "token": "cohere", "startup": True, "website": "https://cohere.com"},
    {"name": "LangChain", "platform": "ashby", "token": "langchain", "startup": True, "website": "https://www.langchain.com"},
    {"name": "Pinecone", "platform": "ashby", "token": "pinecone", "startup": True, "website": "https://www.pinecone.io"},
    {"name": "Chroma", "platform": "ashby", "token": "trychroma", "startup": True, "website": "https://www.trychroma.com"},
    {"name": "IMC Trading", "platform": "greenhouse", "token": "imc", "startup": False, "website": "https://www.imc.com"},
    {"name": "Jane Street", "platform": "greenhouse", "token": "janestreet", "startup": False, "website": "https://www.janestreet.com"},
    {"name": "Squarespace", "platform": "greenhouse", "token": "squarespace", "startup": False, "website": "https://www.squarespace.com"},
    {"name": "Toast", "platform": "greenhouse", "token": "toast", "startup": False, "website": "https://pos.toasttab.com"},
    {"name": "Postman", "platform": "greenhouse", "token": "postman", "startup": True, "website": "https://www.postman.com"},
    {"name": "Docker", "platform": "ashby", "token": "docker", "startup": True, "website": "https://www.docker.com"},
    {"name": "Render", "platform": "ashby", "token": "render", "startup": True, "website": "https://render.com"},
    {"name": "Railway", "platform": "ashby", "token": "railway", "startup": True, "website": "https://railway.com"},
    {"name": "Warp", "platform": "greenhouse", "token": "warp", "startup": True, "website": "https://www.warp.dev"},

    # ── Texas / Arizona coverage (added for the Startups tab region filter) ──
    {"name": "Carvana", "platform": "greenhouse", "token": "carvana", "startup": False, "website": "https://www.carvana.com"},
    {"name": "Axon", "platform": "greenhouse", "token": "axon", "startup": False, "website": "https://www.axon.com"},
    {"name": "CS Disco", "platform": "greenhouse", "token": "disco", "startup": False, "website": "https://www.csdisco.com"},
    {"name": "Bazaarvoice", "platform": "lever", "token": "bazaarvoice", "startup": False, "website": "https://www.bazaarvoice.com"},
    {"name": "Homeward", "platform": "greenhouse", "token": "homeward", "startup": True, "website": "https://www.homeward.com"},
    {"name": "ShiftKey", "platform": "ashby", "token": "shiftkey", "startup": True, "website": "https://www.shiftkey.com"},
    {"name": "AlertMedia", "platform": "greenhouse", "token": "alertmedia", "startup": True, "website": "https://www.alertmedia.com"},
    {"name": "Convey", "platform": "ashby", "token": "convey", "startup": True, "website": "https://www.convey.com"},
    {"name": "Shipwell", "platform": "greenhouse", "token": "shipwell", "startup": True, "website": "https://www.shipwell.com"},
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
# Also covers grad-degree-track titles (Giri is undergrad-only), plus
# founder-level/executive titles common on the YC source (small startups
# often hire "Founding Engineer"/CTO as an experienced-only role).
HARD_AVOID_TITLE_RE = re.compile(
    r"(?i)\b(senior|staff|principal|director|manager|lead|"
    r"founding|chief technology officer|\bcto\b|head of|"
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

MIN_SCORE_TO_INCLUDE = 55  # below this, a posting is dropped entirely (not just low-ranked) — uniform across every source, no exceptions

# The Startups sheet is further narrowed to these regions on top of the
# startup:true tag — Seattle, SF Bay Area, Texas, Arizona, or remote.
STARTUP_TAB_LOCATION_RE = re.compile(
    r"(?i)\b(seattle|bellevue|redmond|kirkland|washington|pullman|"
    r"san francisco|\bsf\b|bay area|oakland|berkeley|san jose|palo alto|"
    r"mountain view|sunnyvale|menlo park|santa clara|fremont|"
    r"texas|austin|dallas|houston|san antonio|fort worth|"
    r"arizona|phoenix|tempe|scottsdale|tucson|"
    r"remote)\b"
)


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
            "website": company["website"],
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
            "website": company["website"],
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
            "website": company["website"],
            "location": location,
            "descriptionText": strip_html(j.get("descriptionHtml") or ""),
            "link": j.get("jobUrl", ""),
            "postedAt": j.get("publishedAt", ""),
        })
    return jobs


FETCHERS = {"greenhouse": fetch_greenhouse, "lever": fetch_lever, "ashby": fetch_ashby}

# ── SimplifyJobs aggregator (layered on top of the per-company APIs) ───────
# Community-maintained, bot-updated every 30-60 min, aggregates postings
# from thousands of companies across many ATS platforms into one JSON file.
# Covers far more companies than we could ever hand-curate — but it has no
# description text, so CATEGORY_TEXT below substitutes for it in scoring.
SIMPLIFY_LISTINGS_URL = (
    "https://raw.githubusercontent.com/SimplifyJobs/"
    "Summer2026-Internships/dev/.github/scripts/listings.json"
)

CATEGORY_TEXT = {
    "Software": "software engineering full stack backend frontend rest api",
    "Software Engineering": "software engineering full stack backend frontend rest api",
    "AI/ML/Data": "artificial intelligence machine learning data analysis data science python",
    "Data Science, AI & Machine Learning": "artificial intelligence machine learning data analysis data science python",
    "Product": "product management business analyst data analytics",
    "Product Management": "product management business analyst data analytics",
    # Hardware/Hardware Engineering/Quant/Quantitative Finance intentionally
    # left out — no skill-boost text, so they fail the skill-overlap gate
    # naturally rather than needing a separate category allowlist.
}

# DOMAIN_EXCLUDE_RE can't catch these — the aggregator has no description
# text, and a defense-hardware company's internship title often doesn't
# literally say "aerospace"/"military" (e.g. "Software Engineer Intern" at
# Anduril). Company-level blocking is the only reliable signal here.
# Used across sources — checked in both fetch_simplify() and fetch_yc(),
# since neither has description text for DOMAIN_EXCLUDE_RE to catch these
# by keyword (a defense-hardware company's internship title often doesn't
# literally say "aerospace"/"military", e.g. Anduril's "Software Engineer
# Intern", or Hop Aero's "Rocket cargo delivery to contested environments"
# one-liner not showing up in the title at all).
DEFENSE_COMPANY_BLOCKLIST = {
    "anduril", "spacex", "space exploration technologies", "lockheed martin",
    "boeing", "raytheon", "rtx", "rtx corporation", "northrop grumman",
    "general dynamics", "l3harris", "leidos", "booz allen hamilton", "saic",
    "bae systems", "textron", "saronic", "shield ai", "epirus", "castelion",
    "firefly aerospace", "rocket lab", "blue origin", "sierra space",
    "redwire", "redwire space", "varda", "varda space", "hadrian",
    "vannevar labs", "applied intuition", "hop aero", "caci", "caci international",
}

# Companies discovered via SimplifyJobs (not in the curated COMPANIES list,
# so isStartup would otherwise default to False) confirmed to be genuine
# startups — this is what actually gets non-SF startups onto the Startups
# tab, since without this, only YC-sourced jobs (which cluster in SF) were
# ever tagged isStartup=True. Grows manually as new ones get spotted;
# defaults stay False for anything not listed here (including large
# companies like Amazon/TikTok/Citadel that also show up via Simplify).
ADDITIONAL_STARTUP_COMPANIES = {
    "docugami",   # Kirkland, WA — AI document startup
    "exowatt",    # Austin, TX — energy hardware startup
}


def fetch_simplify() -> list[dict]:
    r = requests.get(SIMPLIFY_LISTINGS_URL, timeout=30)
    r.raise_for_status()
    entries = r.json()

    known_websites = {c["name"].lower(): c["website"] for c in COMPANIES}
    known_startup = {c["name"].lower(): c["startup"] for c in COMPANIES}
    known_startup.update({name: True for name in ADDITIONAL_STARTUP_COMPANIES})

    jobs = []
    for e in entries:
        if not e.get("active") or not e.get("is_visible", True):
            continue
        if "Bachelor's" not in (e.get("degrees") or []):
            continue  # undergrad-only — same bar as the ATS sources
        title = e.get("title", "")
        company = e.get("company_name", "")
        key = company.lower()
        if key in DEFENSE_COMPANY_BLOCKLIST:
            continue
        posted_ts = e.get("date_posted")
        jobs.append({
            "id": f"simplify:{e['id']}",
            "title": title,
            "companyName": company,
            "isStartup": known_startup.get(key, False),
            "website": known_websites.get(key, e.get("company_url", "")),
            "location": "; ".join(e.get("locations", [])),
            "descriptionText": CATEGORY_TEXT.get(e.get("category", ""), ""),
            "link": e.get("url", ""),
            "postedAt": (
                datetime.fromtimestamp(posted_ts, tz=timezone.utc).isoformat()
                if posted_ts else ""
            ),
        })
    return jobs


# ── Y Combinator's Work at a Startup (layered on top, startup-only) ────────
# No official API — this parses the `data-page` JSON payload Inertia.js
# embeds in the public /jobs pages (robots.txt allows it, no login needed
# to view). More fragile than the JSON-API sources: if YC changes their
# frontend framework or markup, this may need updating.
#
# Every listing here is, by definition, a YC-backed startup — so unlike
# the other sources, jobs found here are always isStartup=True.
#
# Also unlike the other sources: small startups rarely label roles "new
# grad" the way big companies do, so requiring INTERNSHIP_TITLE_RE would
# exclude nearly all full-time YC postings. Instead, scrape_jobs() lets
# anything through here that isn't caught by HARD_AVOID_TITLE_RE (senior/
# staff/founding engineer/CTO/etc. are already excluded there) or
# NON_TECHNICAL_ROLE_RE — covering both internships and next-year
# full-time roles per Giri's request, at the same undergrad-only bar.
YC_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

# Only role categories aligned with the resume — skips Design/Recruiting/
# Sales/Marketing/Legal/Finance/Operations entirely. Value is fallback
# scoring text for that category — used when a posting's own `roleType`
# field is missing, which turns out to be common (e.g. plenty of postings
# on the product-manager category page have roleType: null).
YC_CATEGORY_URLS = {
    "https://www.workatastartup.com/jobs/l/software-engineer":
        "software engineering full stack backend frontend",
    "https://www.workatastartup.com/jobs/l/product-manager":
        "product management roadmap stakeholder data analytics business analyst",
    "https://www.workatastartup.com/jobs/l/science":
        "machine learning artificial intelligence data science research python",
}

# YC's roleType field maps closely to our own skill vocabulary — used as
# scoring text in place of the description text this source doesn't have,
# when present (falls back to the category-level text above otherwise).
YC_ROLE_TYPE_TEXT = {
    "Full stack": "full stack full-stack frontend backend react node",
    "Backend": "backend rest api python",
    "Frontend": "frontend react javascript typescript",
    "Machine learning": "machine learning artificial intelligence deep learning python",
    "Data": "data analysis data science python sql",
    "DevOps": "docker ci/cd aws",
    "Product": "product management roadmap stakeholder data analytics business analyst",
}


def fetch_yc() -> list[dict]:
    known_websites = {c["name"].lower(): c["website"] for c in COMPANIES}
    jobs = []
    seen_ids = set()
    for url, category_text in YC_CATEGORY_URLS.items():
        try:
            r = requests.get(url, headers=YC_HEADERS, timeout=20)
            r.raise_for_status()
            m = re.search(r'data-page="([^"]*)"', r.text)
            if not m:
                raise ValueError("data-page attribute not found — page structure may have changed")
            data = json.loads(html.unescape(m.group(1)))
            for j in data.get("props", {}).get("jobs", []):
                if j["id"] in seen_ids:
                    continue
                seen_ids.add(j["id"])
                key = (j.get("companyName") or "").lower()
                one_liner = j.get("companyOneLiner", "") or ""
                if key in DEFENSE_COMPANY_BLOCKLIST or DOMAIN_EXCLUDE_RE.search(one_liner):
                    continue
                role_text = YC_ROLE_TYPE_TEXT.get(j.get("roleType") or "", "") or category_text
                jobs.append({
                    "id": f"yc:{j['id']}",
                    "title": j.get("title", ""),
                    "companyName": j.get("companyName", ""),
                    "isStartup": True,
                    "website": known_websites.get(key) or (
                        f"https://www.workatastartup.com/companies/{j['companySlug']}"
                        if j.get("companySlug") else ""
                    ),
                    "location": j.get("location", ""),
                    "descriptionText": f"{role_text} {j.get('companyOneLiner', '')}",
                    "link": j.get("applyUrl", ""),
                    "postedAt": "",
                })
        except Exception as e:
            print(f"{'YC ' + url.rsplit('/', 1)[-1]:28s} ({'yc':10s}): FAILED — {e}")
        time.sleep(0.5)
    return jobs


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

    # Same job can be discovered both directly (ATS) and via the aggregator
    # — the direct-ATS copy has real description text, so it wins ties.
    seen_pairs = {(j["companyName"].lower(), j["title"].lower()) for j in all_jobs}

    try:
        simplify_jobs = fetch_simplify()
        internship_jobs = [
            j for j in simplify_jobs
            if INTERNSHIP_TITLE_RE.search(j["title"])
            and not NON_TECHNICAL_ROLE_RE.search(j["title"])
            and (j["companyName"].lower(), j["title"].lower()) not in seen_pairs
        ]
        print(f"{'SimplifyJobs aggregator':28s} ({'simplify':10s}): "
              f"{len(simplify_jobs):4d} total, {len(internship_jobs):3d} internship-tagged")
        all_jobs.extend(internship_jobs)
        seen_pairs.update((j["companyName"].lower(), j["title"].lower()) for j in internship_jobs)
    except Exception as e:
        print(f"{'SimplifyJobs aggregator':28s} ({'simplify':10s}): FAILED — {e}")

    # YC doesn't require INTERNSHIP_TITLE_RE — see fetch_yc()'s docstring
    # comment for why (small startups rarely say "new grad" explicitly).
    try:
        yc_jobs = fetch_yc()
        qualifying_jobs = [
            j for j in yc_jobs
            if not NON_TECHNICAL_ROLE_RE.search(j["title"])
            and (j["companyName"].lower(), j["title"].lower()) not in seen_pairs
        ]
        print(f"{'Y Combinator (WaaS)':28s} ({'yc':10s}): "
              f"{len(yc_jobs):4d} total, {len(qualifying_jobs):3d} undergrad-eligible")
        all_jobs.extend(qualifying_jobs)
    except Exception as e:
        print(f"{'Y Combinator (WaaS)':28s} ({'yc':10s}): FAILED — {e}")

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
    # Exception: an explicit intern/co-op/fellowship/new-grad label overrides
    # this. "manager"/"lead"/"staff" etc. are meant to catch seniority, but
    # "manager" is also just the literal job title for Product roles at any
    # level — "Product Manager Summer 2027 Intern" was getting killed by the
    # bare "manager" match despite plainly self-labeling as an internship.
    hard_hit = HARD_AVOID_TITLE_RE.search(title)
    if hard_hit and not INTERNSHIP_TITLE_RE.search(title):
        return 0, [f"filtered: '{hard_hit.group(0)}' in title"]

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
    elif job["id"].startswith("yc:"):
        # YC full-time postings rarely say "new grad" explicitly even when
        # they're accessible — HARD_AVOID_TITLE_RE already screened out
        # senior/staff/founding-engineer/CTO/lead titles, so what's left
        # gets partial credit rather than zero. Same 55-point bar as every
        # other source applies on top of this — no separate threshold.
        score += 23
        reasons.append("YC full-time (undergrad-eligible)")

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
    # Re-tiered so all 3 labels are still reachable now that
    # MIN_SCORE_TO_INCLUDE=55 means nothing below 55 ever gets written.
    if score >= 80: return "Strong Fit"
    if score >= 65: return "Good Fit"
    return "Possible Fit"


def chance_color(score: int) -> str:
    if score >= 80: return "C6EFCE"   # green
    if score >= 65: return "FFEB9C"   # yellow
    return "FFCC99"                    # orange


# ── Deduplicate against existing sheet ───────────────────────────────────
INTERNSHIPS_SHEET_NAME = "Internships"
JOBS_SHEET_NAME = "Jobs"
STARTUP_SHEET_NAME = "Startups"

# Tab routing: literal internship/co-op/fellowship titles go to the
# Internships tab; everything else that qualified (new-grad, YC full-time)
# goes to Jobs. Startups is a same-day mirror of both, filtered further.
JOB_TYPE_TITLE_RE = re.compile(r"(?i)\b(intern(ship)?s?|co-?op|fellow(ship)?s?)\b")


def is_internship_type(title: str) -> bool:
    return bool(JOB_TYPE_TITLE_RE.search(title))


def load_existing_ids(path: str) -> set:
    if not os.path.exists(path):
        return set()
    try:
        wb = openpyxl.load_workbook(path)
        ids = set()
        for name in (INTERNSHIPS_SHEET_NAME, JOBS_SHEET_NAME):
            if name in wb.sheetnames:
                ids.update(str(row[0]) for row in wb[name].iter_rows(min_row=2, values_only=True) if row[0])
        if not ids and wb.sheetnames:
            # Backward compat with the old single-sheet layout
            ids.update(str(row[0]) for row in wb.active.iter_rows(min_row=2, values_only=True) if row[0])
        return ids
    except Exception:
        return set()


# ── Outreach drafts ────────────────────────────────────────────────────────
# Message text and manual-search links only — nothing here scrapes LinkedIn
# or sends anything. Finding the actual person and hitting send is on Giri.
def _matched_skills(job: dict) -> list[str]:
    full_text = f"{(job.get('title') or '').lower()} {(job.get('descriptionText') or '').lower()}"
    return [s for s, pat in SKILL_PATTERNS if pat.search(full_text)]


def _pick_highlight(matched_skills: list[str]) -> dict:
    for category in ("ml", "web", "data"):
        if any(s in HIGHLIGHTS[category]["keywords"] for s in matched_skills):
            return HIGHLIGHTS[category]
    return HIGHLIGHTS["default"]


# WSU alumni search — kept as the single link instead of two separate ones
# (alumni + recruiter). A fellow Coug is a far better warm-intro target
# than a random recruiter, so this is "the better option" between them.
def linkedin_search_link(company: str) -> str:
    query = f"{company} {CONTACT['school']}"
    return f"https://www.linkedin.com/search/results/people/?keywords={quote(query)}"


def generate_linkedin_note(job: dict) -> str:
    highlight = _pick_highlight(_matched_skills(job))
    title, company = job.get("title", "this role"), job.get("companyName", "")
    note = (
        f"Hi! I'm a CS student at {CONTACT['school']} and saw {company} is hiring for "
        f"{title}. {highlight['short']} — would love to connect and hear more about the team."
    )
    if len(note) > 300:
        note = (
            f"Hi! WSU CS student here, saw {company} is hiring for {title}. "
            f"{highlight['short']} — would love to connect."
        )[:300]
    return note


def generate_email(job: dict) -> str:
    highlight = _pick_highlight(_matched_skills(job))
    title, company = job.get("title", "this role"), job.get("companyName", "")
    subject = f"Interested in the {title} role at {company}"
    body = (
        f"Subject: {subject}\n\n"
        f"Hi,\n\n"
        f"My name is {CONTACT['name']}, and I'm a Computer Science student at {CONTACT['school']} "
        f"(graduating {CONTACT['grad']}). I came across the {title} opening at {company} and wanted "
        f"to reach out directly — I'm genuinely interested in the role and think my background lines up well.\n\n"
        f"{highlight['email']}\n\n"
        f"I'd love to learn more about the team and what you're looking for — happy to send my resume "
        f"or hop on a quick call if you're open to it.\n\n"
        f"Thanks so much for your time,\n"
        f"{CONTACT['name']}\n"
        f"{CONTACT['phone']} | {CONTACT['email']} | {CONTACT['github']}"
    )
    return body


# ── Excel builder ─────────────────────────────────────────────────────────
HEADERS = [
    "Job ID", "Date Found", "Title", "Company", "Company Website", "Location",
    "Match %", "Fit", "Source", "Why It Matched", "Link",
    "LinkedIn Search", "LinkedIn Note Draft", "Email Draft",
    "Status", "Notes",
]

COL_WIDTHS = [24, 12, 34, 24, 22, 20, 10, 13, 10, 40, 14, 16, 50, 50, 14, 24]

NAVY = "1A3A5C"
WHITE = "FFFFFF"
LIGHT = "F2F5F9"
BORDER_COLOR = "CCCCCC"


def thin_border():
    s = Side(style="thin", color=BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)


def write_excel(jobs_scored: list[dict], path: str) -> list[tuple[dict, int]]:
    """Returns the (job, score) pairs actually written — used to build the
    Discord notification so it only reports what's genuinely new today.
    Routes each job to Internships or Jobs by title; Startups is a
    same-day mirror of both, filtered further by startup tag + region."""
    existing_ids = load_existing_ids(path)
    added = []

    if os.path.exists(path):
        wb = openpyxl.load_workbook(path)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    sheets = {}
    for name in (INTERNSHIPS_SHEET_NAME, JOBS_SHEET_NAME, STARTUP_SHEET_NAME):
        if name in wb.sheetnames:
            sheets[name] = wb[name]
        else:
            sheets[name] = wb.create_sheet(name)
            _write_header(sheets[name])
    ws_internships, ws_jobs, ws_startup = (
        sheets[INTERNSHIPS_SHEET_NAME], sheets[JOBS_SHEET_NAME], sheets[STARTUP_SHEET_NAME]
    )

    startup_count = 0
    for job in jobs_scored:
        if job["id"] in existing_ids:
            continue
        target_ws = ws_internships if is_internship_type(job["title"]) else ws_jobs
        score = _append_row(target_ws, job)
        if score is not None:
            added.append((job, score))
            is_target_region = STARTUP_TAB_LOCATION_RE.search((job.get("location") or "").lower())
            if job.get("isStartup") and is_target_region:
                _append_row(ws_startup, job)
                startup_count += 1

    print(f"Added {len(added)} new postings "
          f"({startup_count} startup, Seattle/SF Bay/TX/AZ/remote) to the sheet.")

    for ws in (ws_internships, ws_jobs, ws_startup):
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
    company = job.get("companyName", "")

    values = [
        job["id"],
        datetime.today().strftime("%Y-%m-%d"),
        job.get("title", ""),
        company,
        job.get("website", ""),
        job.get("location", ""),
        score,
        chance_label(score),
        source,
        ", ".join(reasons[:4]),
        job.get("link", ""),
        linkedin_search_link(company),
        generate_linkedin_note(job),
        generate_email(job),
        "Not Applied",
        "",
    ]

    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = Font(name="Arial", size=9)
        cell.alignment = Alignment(vertical="center", wrap_text=(col in [3, 10, 13, 14]))
        cell.border = thin_border()
        if col not in [7, 8]:
            cell.fill = PatternFill("solid", fgColor=fill_color)

    pct_cell = ws.cell(row=row, column=7)
    pct_cell.value = f"{score}%"
    pct_cell.font = Font(name="Arial", size=9, bold=True)
    pct_cell.fill = PatternFill("solid", fgColor=chance_color(score))
    pct_cell.alignment = Alignment(horizontal="center", vertical="center")

    ch_cell = ws.cell(row=row, column=8)
    ch_cell.fill = PatternFill("solid", fgColor=chance_color(score))
    ch_cell.alignment = Alignment(horizontal="center", vertical="center")

    website = values[4]
    if website:
        site_cell = ws.cell(row=row, column=5, value="Website →")
        site_cell.hyperlink = website
        site_cell.font = Font(name="Arial", size=9, color="185FA5", underline="single")
        site_cell.alignment = Alignment(horizontal="center", vertical="center")

    link = values[10]
    if link:
        link_cell = ws.cell(row=row, column=11, value="Apply →")
        link_cell.hyperlink = link
        link_cell.font = Font(name="Arial", size=9, color="185FA5", underline="single")
        link_cell.alignment = Alignment(horizontal="center", vertical="center")

    search_cell = ws.cell(row=row, column=12, value="LinkedIn Search →")
    search_cell.hyperlink = values[11]
    search_cell.font = Font(name="Arial", size=9, color="185FA5", underline="single")
    search_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[row].height = 60
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
