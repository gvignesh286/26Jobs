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
    # handled separately by SENIORITY_AVOID_RE/GRAD_DEGREE_AVOID_RE (a hard drop, not a penalty).
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

    # ── Added: 100+ company expansion across CS-adjacent sectors ──
    # (fintech, healthtech, insurtech, AI/ML, cybersecurity, dev tools,
    # gaming, logistics, climate tech, consumer, edtech, quant trading,
    # crypto, biotech, autonomous vehicles, legal tech, business/productivity
    # SaaS) — each token verified against the live Greenhouse/Lever/Ashby
    # API before being added; a handful of generic-word token guesses that
    # resolved to a DIFFERENT real company of the same name were caught by
    # sampling actual job titles/locations and excluded.

    # ── AI / ML ──
    {"name": "ElevenLabs", "platform": "ashby", "token": "elevenlabs", "startup": True, "website": "https://elevenlabs.io"},
    {"name": "Glean", "platform": "greenhouse", "token": "gleanwork", "startup": True, "website": "https://www.glean.com"},
    {"name": "Inflection AI", "platform": "greenhouse", "token": "inflectionai", "startup": True, "website": "https://inflection.ai"},
    {"name": "Mistral AI", "platform": "lever", "token": "mistral", "startup": True, "website": "https://mistral.ai"},
    {"name": "Speak", "platform": "ashby", "token": "speak", "startup": True, "website": "https://www.speak.com"},
    {"name": "Synthesia", "platform": "ashby", "token": "synthesia", "startup": True, "website": "https://www.synthesia.io"},
    {"name": "World Labs", "platform": "greenhouse", "token": "worldlabs", "startup": True, "website": "https://www.worldlabs.ai"},
    {"name": "Writer", "platform": "ashby", "token": "writer", "startup": True, "website": "https://writer.com"},

    # ── Autonomous vehicles ──
    {"name": "Lucid Motors", "platform": "greenhouse", "token": "lucidmotors", "startup": False, "website": "https://www.lucidmotors.com"},
    {"name": "Waymo", "platform": "greenhouse", "token": "waymo", "startup": False, "website": "https://waymo.com"},
    {"name": "Zoox", "platform": "lever", "token": "zoox", "startup": False, "website": "https://zoox.com"},

    # ── Biotech / life sciences ──
    {"name": "Benchling", "platform": "ashby", "token": "benchling", "startup": True, "website": "https://www.benchling.com"},
    {"name": "Freenome", "platform": "greenhouse", "token": "freenome", "startup": True, "website": "https://www.freenome.com"},
    {"name": "Ginkgo Bioworks", "platform": "greenhouse", "token": "ginkgobioworks", "startup": False, "website": "https://www.ginkgobioworks.com"},
    {"name": "Grail", "platform": "lever", "token": "grailbio", "startup": False, "website": "https://grail.com"},
    {"name": "Insitro", "platform": "ashby", "token": "insitro", "startup": True, "website": "https://www.insitro.com"},

    # ── Business / productivity SaaS ──
    {"name": "AngelList", "platform": "lever", "token": "angellist", "startup": True, "website": "https://www.angellist.com"},
    {"name": "ClickUp", "platform": "ashby", "token": "clickup", "startup": True, "website": "https://clickup.com"},
    {"name": "Deel", "platform": "ashby", "token": "deel", "startup": True, "website": "https://www.deel.com"},
    {"name": "Freshworks", "platform": "lever", "token": "freshworks", "startup": False, "website": "https://www.freshworks.com"},
    {"name": "Front", "platform": "ashby", "token": "frontapp", "startup": True, "website": "https://front.com"},
    {"name": "Gusto", "platform": "greenhouse", "token": "gusto", "startup": True, "website": "https://gusto.com"},
    {"name": "HubSpot", "platform": "greenhouse", "token": "hubspot", "startup": False, "website": "https://www.hubspot.com"},
    {"name": "Instawork", "platform": "greenhouse", "token": "instawork", "startup": True, "website": "https://www.instawork.com"},
    {"name": "Intercom", "platform": "greenhouse", "token": "intercom", "startup": True, "website": "https://www.intercom.com"},
    {"name": "Loom", "platform": "ashby", "token": "loom", "startup": True, "website": "https://www.loom.com"},
    {"name": "Middesk", "platform": "ashby", "token": "middesk", "startup": True, "website": "https://www.middesk.com"},
    {"name": "Olo", "platform": "lever", "token": "olo", "startup": False, "website": "https://www.olo.com"},
    {"name": "SpotOn", "platform": "ashby", "token": "spoton", "startup": True, "website": "https://spoton.com"},
    {"name": "Superhuman", "platform": "ashby", "token": "superhuman", "startup": True, "website": "https://superhuman.com"},
    {"name": "Wonolo", "platform": "lever", "token": "wonolo", "startup": True, "website": "https://www.wonolo.com"},
    {"name": "Y Combinator", "platform": "ashby", "token": "ycombinator", "startup": True, "website": "https://www.ycombinator.com"},
    {"name": "ezCater", "platform": "lever", "token": "ezcater", "startup": True, "website": "https://www.ezcater.com"},

    # ── Climate / energy tech ──
    {"name": "Aurora Solar", "platform": "ashby", "token": "aurorasolar", "startup": True, "website": "https://www.aurorasolar.com"},
    {"name": "Redwood Materials", "platform": "greenhouse", "token": "redwoodmaterials", "startup": True, "website": "https://www.redwoodmaterials.com"},
    {"name": "Span", "platform": "ashby", "token": "span", "startup": True, "website": "https://www.span.io"},

    # ── Consumer / retail ──
    {"name": "Airbnb", "platform": "greenhouse", "token": "airbnb", "startup": False, "website": "https://www.airbnb.com"},
    {"name": "Away", "platform": "ashby", "token": "away", "startup": True, "website": "https://www.awaytravel.com"},
    {"name": "Calm", "platform": "greenhouse", "token": "calm", "startup": True, "website": "https://www.calm.com"},
    {"name": "Cameo", "platform": "greenhouse", "token": "cameo", "startup": True, "website": "https://www.cameo.com"},
    {"name": "ClassPass", "platform": "greenhouse", "token": "classpass", "startup": True, "website": "https://classpass.com"},
    {"name": "Glossier", "platform": "greenhouse", "token": "glossier", "startup": True, "website": "https://www.glossier.com"},
    {"name": "Life360", "platform": "greenhouse", "token": "life360", "startup": False, "website": "https://www.life360.com"},
    {"name": "Lyft", "platform": "greenhouse", "token": "lyft", "startup": False, "website": "https://www.lyft.com"},
    {"name": "Nextdoor", "platform": "greenhouse", "token": "nextdoor", "startup": False, "website": "https://about.nextdoor.com"},
    {"name": "Oura", "platform": "greenhouse", "token": "oura", "startup": True, "website": "https://ouraring.com"},
    {"name": "Patreon", "platform": "ashby", "token": "patreon", "startup": True, "website": "https://www.patreon.com"},
    {"name": "Peloton", "platform": "greenhouse", "token": "peloton", "startup": False, "website": "https://www.onepeloton.com"},
    {"name": "Poshmark", "platform": "greenhouse", "token": "poshmark", "startup": False, "website": "https://poshmark.com"},
    {"name": "Rent the Runway", "platform": "greenhouse", "token": "renttherunway", "startup": False, "website": "https://www.renttherunway.com"},
    {"name": "Spotify", "platform": "lever", "token": "spotify", "startup": False, "website": "https://www.spotify.com"},
    {"name": "Strava", "platform": "ashby", "token": "strava", "startup": True, "website": "https://www.strava.com"},
    {"name": "TaskRabbit", "platform": "greenhouse", "token": "taskrabbit", "startup": False, "website": "https://www.taskrabbit.com"},
    {"name": "Thumbtack", "platform": "ashby", "token": "thumbtack", "startup": True, "website": "https://www.thumbtack.com"},

    # ── Crypto / Web3 ──
    {"name": "Alchemy", "platform": "ashby", "token": "alchemy", "startup": True, "website": "https://www.alchemy.com"},
    {"name": "BitGo", "platform": "greenhouse", "token": "bitgo", "startup": True, "website": "https://www.bitgo.com"},
    {"name": "Consensys", "platform": "greenhouse", "token": "consensys", "startup": True, "website": "https://consensys.io"},
    {"name": "Kraken", "platform": "lever", "token": "kraken", "startup": True, "website": "https://www.kraken.com"},
    {"name": "MoonPay", "platform": "lever", "token": "moonpay", "startup": True, "website": "https://www.moonpay.com"},
    {"name": "OpenSea", "platform": "ashby", "token": "opensea", "startup": True, "website": "https://opensea.io"},
    {"name": "Ripple", "platform": "greenhouse", "token": "ripple", "startup": True, "website": "https://ripple.com"},
    {"name": "Solana Labs", "platform": "ashby", "token": "solanalabs", "startup": True, "website": "https://solanalabs.com"},

    # ── Cybersecurity ──
    {"name": "Abnormal Security", "platform": "greenhouse", "token": "abnormalsecurity", "startup": True, "website": "https://abnormalsecurity.com"},
    {"name": "Huntress", "platform": "greenhouse", "token": "huntress", "startup": True, "website": "https://www.huntress.com"},
    {"name": "Netskope", "platform": "greenhouse", "token": "netskope", "startup": True, "website": "https://www.netskope.com"},
    {"name": "Rubrik", "platform": "greenhouse", "token": "rubrik", "startup": False, "website": "https://www.rubrik.com"},
    {"name": "Semgrep", "platform": "ashby", "token": "semgrep", "startup": True, "website": "https://semgrep.dev"},
    {"name": "Snyk", "platform": "ashby", "token": "snyk", "startup": True, "website": "https://snyk.io"},
    {"name": "Tanium", "platform": "greenhouse", "token": "tanium", "startup": True, "website": "https://www.tanium.com"},
    {"name": "Wiz", "platform": "ashby", "token": "wiz", "startup": True, "website": "https://www.wiz.io"},
    {"name": "Zscaler", "platform": "greenhouse", "token": "zscaler", "startup": False, "website": "https://www.zscaler.com"},

    # ── Dev tools / data infra ──
    {"name": "CircleCI", "platform": "greenhouse", "token": "circleci", "startup": True, "website": "https://circleci.com"},
    {"name": "Cribl", "platform": "greenhouse", "token": "cribl", "startup": True, "website": "https://cribl.io"},
    {"name": "Fivetran", "platform": "greenhouse", "token": "fivetran", "startup": True, "website": "https://www.fivetran.com"},
    {"name": "Grafana Labs", "platform": "greenhouse", "token": "grafanalabs", "startup": True, "website": "https://grafana.com"},
    {"name": "Hightouch", "platform": "greenhouse", "token": "hightouch", "startup": True, "website": "https://hightouch.com"},
    {"name": "Monte Carlo", "platform": "ashby", "token": "montecarlodata", "startup": True, "website": "https://www.montecarlodata.com"},
    {"name": "New Relic", "platform": "greenhouse", "token": "newrelic", "startup": False, "website": "https://newrelic.com"},
    {"name": "Nutanix", "platform": "ashby", "token": "nutanix", "startup": False, "website": "https://www.nutanix.com"},
    {"name": "Sentry", "platform": "ashby", "token": "sentry", "startup": True, "website": "https://sentry.io"},

    # ── EdTech ──
    {"name": "Course Hero", "platform": "greenhouse", "token": "coursehero", "startup": True, "website": "https://www.coursehero.com"},
    {"name": "Coursera", "platform": "greenhouse", "token": "coursera", "startup": False, "website": "https://www.coursera.org"},
    {"name": "Guild Education", "platform": "greenhouse", "token": "guild", "startup": True, "website": "https://www.guild.com"},
    {"name": "Handshake", "platform": "ashby", "token": "handshake", "startup": True, "website": "https://joinhandshake.com"},
    {"name": "Outschool", "platform": "greenhouse", "token": "outschool", "startup": True, "website": "https://outschool.com"},

    # ── Fintech ──
    {"name": "Acorns", "platform": "ashby", "token": "acorns", "startup": True, "website": "https://www.acorns.com"},
    {"name": "Bill.com", "platform": "greenhouse", "token": "billcom", "startup": False, "website": "https://www.bill.com"},
    {"name": "Brex", "platform": "greenhouse", "token": "brex", "startup": True, "website": "https://www.brex.com"},
    {"name": "Dave", "platform": "ashby", "token": "dave", "startup": False, "website": "https://www.dave.com"},
    {"name": "Highbeam", "platform": "ashby", "token": "highbeam", "startup": True, "website": "https://www.highbeam.co"},
    {"name": "Marqeta", "platform": "greenhouse", "token": "marqeta", "startup": False, "website": "https://www.marqeta.com"},
    {"name": "NerdWallet", "platform": "ashby", "token": "nerdwallet", "startup": False, "website": "https://www.nerdwallet.com"},
    {"name": "Novo", "platform": "ashby", "token": "novo", "startup": True, "website": "https://www.novo.co"},
    {"name": "Public.com", "platform": "greenhouse", "token": "public", "startup": True, "website": "https://public.com"},
    {"name": "Synctera", "platform": "ashby", "token": "synctera", "startup": True, "website": "https://synctera.com"},
    {"name": "Treasury Prime", "platform": "greenhouse", "token": "treasuryprime", "startup": True, "website": "https://www.treasuryprime.com"},

    # ── Gaming ──
    {"name": "Epic Games", "platform": "greenhouse", "token": "epicgames", "startup": True, "website": "https://www.epicgames.com"},
    {"name": "Niantic", "platform": "ashby", "token": "niantic", "startup": True, "website": "https://nianticlabs.com"},
    {"name": "Riot Games", "platform": "greenhouse", "token": "riotgames", "startup": True, "website": "https://www.riotgames.com"},
    {"name": "Rockstar Games", "platform": "greenhouse", "token": "rockstargames", "startup": False, "website": "https://www.rockstargames.com"},
    {"name": "Scopely", "platform": "greenhouse", "token": "scopely", "startup": True, "website": "https://www.scopely.com"},

    # ── Healthtech ──
    {"name": "Abridge", "platform": "ashby", "token": "abridge", "startup": True, "website": "https://www.abridge.com"},
    {"name": "Aledade", "platform": "lever", "token": "aledade", "startup": True, "website": "https://www.aledade.com"},
    {"name": "Ambience Healthcare", "platform": "ashby", "token": "ambiencehealthcare", "startup": True, "website": "https://www.ambiencehealthcare.com"},
    {"name": "Arcadia", "platform": "lever", "token": "arcadia", "startup": True, "website": "https://arcadia.io"},
    {"name": "Butterfly Network", "platform": "greenhouse", "token": "butterflynetwork", "startup": False, "website": "https://www.butterflynetwork.com"},
    {"name": "Carbon Health", "platform": "lever", "token": "carbonhealth", "startup": True, "website": "https://carbonhealth.com"},
    {"name": "Cerebral", "platform": "greenhouse", "token": "cerebral", "startup": True, "website": "https://cerebral.com"},
    {"name": "Cityblock Health", "platform": "ashby", "token": "cityblock", "startup": True, "website": "https://www.cityblock.com"},
    {"name": "Clover Health", "platform": "greenhouse", "token": "cloverhealth", "startup": False, "website": "https://www.cloverhealth.com"},
    {"name": "Cohere Health", "platform": "greenhouse", "token": "coherehealth", "startup": True, "website": "https://www.coherehealth.com"},
    {"name": "Commure", "platform": "ashby", "token": "commure", "startup": True, "website": "https://www.commure.com"},
    {"name": "Curai Health", "platform": "lever", "token": "curai", "startup": True, "website": "https://curaihealth.com"},
    {"name": "Doximity", "platform": "greenhouse", "token": "doximity", "startup": False, "website": "https://www.doximity.com"},
    {"name": "Found", "platform": "greenhouse", "token": "found", "startup": True, "website": "https://www.joinfound.com"},
    {"name": "Headway", "platform": "ashby", "token": "headway", "startup": True, "website": "https://headway.co"},
    {"name": "Included Health", "platform": "lever", "token": "includedhealth", "startup": True, "website": "https://www.includedhealth.com"},
    {"name": "Komodo Health", "platform": "greenhouse", "token": "komodohealth", "startup": True, "website": "https://www.komodohealth.com"},
    {"name": "Lyra Health", "platform": "lever", "token": "lyrahealth", "startup": True, "website": "https://www.lyrahealth.com"},
    {"name": "Maven Clinic", "platform": "greenhouse", "token": "mavenclinic", "startup": True, "website": "https://www.mavenclinic.com"},
    {"name": "Modern Health", "platform": "greenhouse", "token": "modernhealth", "startup": True, "website": "https://www.modernhealth.com"},
    {"name": "Notable Health", "platform": "ashby", "token": "notable", "startup": True, "website": "https://www.notablehealth.com"},
    {"name": "Nuna", "platform": "ashby", "token": "nuna", "startup": True, "website": "https://www.nuna.com"},
    {"name": "One Medical", "platform": "greenhouse", "token": "onemedical", "startup": False, "website": "https://www.onemedical.com"},
    {"name": "Overjet", "platform": "ashby", "token": "overjet", "startup": True, "website": "https://www.overjet.com"},
    {"name": "Parsley Health", "platform": "greenhouse", "token": "parsleyhealth", "startup": True, "website": "https://www.parsleyhealth.com"},
    {"name": "PathAI", "platform": "greenhouse", "token": "pathai", "startup": True, "website": "https://www.pathai.com"},
    {"name": "Ro", "platform": "lever", "token": "ro", "startup": True, "website": "https://ro.co"},
    {"name": "Sidecar Health", "platform": "greenhouse", "token": "sidecarhealth", "startup": True, "website": "https://sidecarhealth.com"},
    {"name": "Suki AI", "platform": "greenhouse", "token": "suki", "startup": True, "website": "https://www.suki.ai"},
    {"name": "Sword Health", "platform": "greenhouse", "token": "swordhealth", "startup": True, "website": "https://www.swordhealth.com"},
    {"name": "Talkspace", "platform": "greenhouse", "token": "talkspace", "startup": False, "website": "https://www.talkspace.com"},
    {"name": "Truveta", "platform": "greenhouse", "token": "truveta", "startup": True, "website": "https://truveta.com"},
    {"name": "Waymark", "platform": "greenhouse", "token": "waymark", "startup": True, "website": "https://waymarkcare.com"},
    {"name": "Zocdoc", "platform": "greenhouse", "token": "zocdoc", "startup": True, "website": "https://www.zocdoc.com"},

    # ── Insurtech ──
    {"name": "At-Bay", "platform": "greenhouse", "token": "atbay", "startup": True, "website": "https://www.at-bay.com"},
    {"name": "Bestow", "platform": "ashby", "token": "bestow", "startup": True, "website": "https://www.bestow.com"},
    {"name": "Branch Insurance", "platform": "ashby", "token": "branchinsurance", "startup": True, "website": "https://www.ourbranch.com"},
    {"name": "Coalition", "platform": "greenhouse", "token": "coalition", "startup": True, "website": "https://www.coalitioninc.com"},
    {"name": "Ethos Life", "platform": "greenhouse", "token": "ethoslife", "startup": True, "website": "https://www.ethoslife.com"},
    {"name": "Kin Insurance", "platform": "ashby", "token": "kin", "startup": True, "website": "https://www.kin.com"},
    {"name": "Lemonade", "platform": "ashby", "token": "lemonade", "startup": False, "website": "https://www.lemonade.com"},
    {"name": "Openly", "platform": "ashby", "token": "openly", "startup": True, "website": "https://www.openly.com"},
    {"name": "Pie Insurance", "platform": "greenhouse", "token": "pieinsurance", "startup": True, "website": "https://pieinsurance.com"},

    # ── Legal tech ──
    {"name": "Everlaw", "platform": "greenhouse", "token": "everlaw", "startup": True, "website": "https://www.everlaw.com"},

    # ── Logistics / supply chain ──
    {"name": "Bringg", "platform": "greenhouse", "token": "bringg", "startup": True, "website": "https://www.bringg.com"},
    {"name": "DoorDash", "platform": "greenhouse", "token": "doordashusa", "startup": False, "website": "https://www.doordash.com"},
    {"name": "Flexe", "platform": "greenhouse", "token": "flexe", "startup": True, "website": "https://www.flexe.com"},
    {"name": "FourKites", "platform": "greenhouse", "token": "fourkites", "startup": True, "website": "https://www.fourkites.com"},
    {"name": "Gopuff", "platform": "lever", "token": "gopuff", "startup": True, "website": "https://gopuff.com"},
    {"name": "Loadsmart", "platform": "lever", "token": "loadsmart", "startup": True, "website": "https://www.loadsmart.com"},
    {"name": "Uber Freight", "platform": "greenhouse", "token": "uberfreight", "startup": False, "website": "https://www.uberfreight.com"},
    {"name": "Zipline", "platform": "greenhouse", "token": "flyzipline", "startup": True, "website": "https://www.flyzipline.com"},
    {"name": "project44", "platform": "greenhouse", "token": "project44", "startup": True, "website": "https://www.project44.com"},

    # ── Marketing tech ──
    {"name": "Attentive", "platform": "greenhouse", "token": "attentive", "startup": True, "website": "https://www.attentive.com"},

    # ── Quant trading / prop trading ──
    {"name": "Akuna Capital", "platform": "greenhouse", "token": "akunacapital", "startup": False, "website": "https://akunacapital.com"},
    {"name": "Belvedere Trading", "platform": "lever", "token": "belvederetrading", "startup": False, "website": "https://www.belvederetrading.com"},
    {"name": "Hudson River Trading", "platform": "ashby", "token": "hrt", "startup": False, "website": "https://www.hudsonrivertrading.com"},
    {"name": "Jump Trading", "platform": "greenhouse", "token": "jumptrading", "startup": False, "website": "https://www.jumptrading.com"},
    {"name": "Optiver", "platform": "greenhouse", "token": "optiver", "startup": False, "website": "https://optiver.com"},
    {"name": "Point72", "platform": "greenhouse", "token": "point72", "startup": False, "website": "https://www.point72.com"},
    {"name": "Tower Research Capital", "platform": "greenhouse", "token": "towerresearchcapital", "startup": False, "website": "https://www.tower-research.com"},

    # ── Space ──
    {"name": "Astranis", "platform": "greenhouse", "token": "astranis", "startup": True, "website": "https://www.astranis.com"},
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
    r"medical|clinical|"
    # Clinical/lab provider & scientist titles - a title-level backstop for
    # when roleType-based exclusion (see YC_EXCLUDED_ROLE_TYPES) isn't
    # available or reliable. Found via real postings that slipped through:
    # "Licensed Mental Health Providers", "Psychiatric Nurse Practitioner
    # (PMHNP)", "Formulation Scientist", "Telemedicine Specialist".
    r"nurse practitioner|physician|pharmacist|\bpharmacy\b|dietitian|"
    r"psychiatric|mental health provider|licensed (clinical|therapist)|"
    r"telemedicine specialist|formulation scientist|"
    r"clinical (researcher|lead)|diagnostics clinical)\b"
)

# Hard title exclusions, split into two categories that get treated
# differently in score_job():
#
# SENIORITY — an explicit intern/co-op/fellowship/new-grad label in the
# title overrides this (e.g. "Product Manager Summer 2027 Intern" — the
# base title "Product Manager" isn't a seniority signal by itself once
# it's explicitly labeled an internship). Also covers founder-level/
# executive titles common on the YC source (small startups often hire
# "Founding Engineer"/CTO as an experienced-only role).
SENIORITY_AVOID_RE = re.compile(
    r"(?i)\b(senior|sr\.?|staff|principal|director|manager|lead|"
    r"founding|chief technology officer|\bcto\b|head of|"
    r"postdoc(toral)?|research scientist)\b"
)

# GRAD-DEGREE — never overridable, regardless of "intern" appearing in the
# same title. A posting can be both an internship AND require a graduate
# degree — e.g. a real one found on the UW source: "2026 Fall Applied
# Science Internship – ... PhD Student Science Recruiting". Undergrad-only
# eligibility is an absolute requirement, not a seniority judgment call.
# "ph\.d\." with a trailing \b fails on real-world titles like "(Ph.D.)" —
# both the matched "." and the following ")" are non-word chars, so no
# word-boundary transition exists there and \b never fires. Found via a
# real UW posting ("Statistical Scientist (Ph.D.)") slipping through with
# only a soft penalty instead of the intended hard exclusion. Uses the same
# lookaround boundary style as skill_pattern() for the phd variants.
GRAD_DEGREE_AVOID_RE = re.compile(
    r"(?i)(?<![a-z0-9])(phd|ph\.d\.?)(?![a-z0-9])|"
    r"\b(phd student|graduate program|"
    r"master'?s?\s+student|doctoral( student)?|graduate researcher|mba)\b"
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

# Underclassman eligibility signals. Giri is now a graduating senior (grad
# May 2027) targeting New Grad roles — an internship explicitly gated to
# students with 2+ years of school left (freshmen/sophomores/juniors) isn't
# a fit even if it otherwise looks like a strong internship match. Checked
# against description text (this eligibility language is essentially never
# in the title itself). Internships open to Giri's own class, or with no
# class-year restriction stated, are NOT caught by this — only explicit
# underclassman-targeting language is.
UNDERCLASSMAN_ELIGIBILITY_RE = re.compile(
    r"(?i)\b(rising sophomores?|rising juniors?|rising seniors?|freshm(a|e)n|"
    r"first-year students?|second-year students?|third-year students?|"
    r"sophomore year|penultimate year|"
    r"at least one more (year|semester) (of school|remaining)|"
    r"must (have|complete) (at least )?one more year|"
    r"not (in your|graduating in your) final year)\b"
)

# "Class of 20XX" eligibility requirements naming a grad year later than
# Giri's own (see CONTACT["grad"]) also signal the posting wants someone
# with more school left. Extracted and compared numerically against
# GRAD_YEAR rather than hardcoding specific years, so this doesn't quietly
# go stale as time passes.
CLASS_OF_YEAR_RE = re.compile(r"(?i)\bclass of (20\d{2})\b")
GRAD_YEAR = int(re.search(r"\d{4}", CONTACT["grad"]).group())


def targets_underclassman(text: str) -> str | None:
    """Returns the matched phrase if `text` targets a student with more
    than one year of school left after Giri's own graduation, else None."""
    hit = UNDERCLASSMAN_ELIGIBILITY_RE.search(text)
    if hit:
        return hit.group(0)
    for m in CLASS_OF_YEAR_RE.finditer(text):
        if int(m.group(1)) > GRAD_YEAR:
            return m.group(0)
    return None


# Hard cutoff — a posting older than this is dropped entirely, not just
# down-scored. Only applies where a real date is known (see score_job()) —
# sources with no postedAt (YC, UW) can't be proven stale, so they aren't
# dropped by this; same reasoning as everywhere else a signal is missing
# rather than negative.
MAX_POSTING_AGE_DAYS = 7

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
# anything through here that isn't caught by SENIORITY_AVOID_RE/GRAD_DEGREE_AVOID_RE (senior/
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

# roleType values that are explicitly non-technical/wrong-domain for a CS
# candidate. Found via a real bug: the "/science" category page returns
# wet-lab/clinical roles too (Formulation Scientist -> roleType
# "Biotechnology", Licensed Mental Health Providers / Psychiatric Nurse
# Practitioner -> roleType "Healthcare") - none of those are in
# YC_ROLE_TYPE_TEXT, so they fell through to the URL-level "machine
# learning artificial intelligence data science research python" fallback
# meant for postings with NO roleType at all, not ones that explicitly say
# otherwise. That fake description then passed the skill-match gate and
# scored 60% in production. Excluded outright here rather than trusting
# downstream filters, since the fake description was defeating them.
YC_EXCLUDED_ROLE_TYPES = {
    "Healthcare", "Biotechnology", "Biology", "Chemistry", "Immunology",
    "Oncology", "Laboratory",
}

# YC's location field is a flat string, not structured data - checked
# separately from NON_US_LOCATION_RE (built for spelled-out city/country
# names) because YC formats non-US locations with 2-3 letter codes instead
# ("Gurugram, HR, IN / Remote (IN)"), which that regex doesn't catch. Also
# a real bug found in production: PROFILE["preferred_locations"] contains
# "remote", matched as a bare substring - "Remote (IN)" contains "remote"
# and was scoring +20 location points as if it were US-remote.
US_STATE_CODES = {
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID",
    "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS",
    "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK",
    "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV",
    "WI", "WY", "DC", "PR",
}


def _yc_location_is_non_us(location: str) -> bool:
    loc = (location or "").strip()
    if not loc:
        return False
    # "... / Remote (IN)" - explicit non-US remote marker.
    m = re.search(r"remote\s*\(([a-z]{2,3})\)", loc.lower())
    if m and m.group(1).upper() not in ("US", "USA"):
        return True
    primary = loc.split("/")[0].strip()
    parts = [p.strip() for p in primary.split(",")]
    if len(parts) >= 3:
        # YC's 3-part format is "City, State/Region, CountryCode". For
        # genuine US postings the country segment is literally "US"/"USA"
        # (e.g. "San Francisco, CA, US"), never a repeated/different state
        # code - checking the last segment against the state-code
        # whitelist would wrongly accept "Bengaluru, KA, IN" since IN is
        # *also* Indiana's code. A real gap found in production: this
        # exact string slipped past an earlier version of this check that
        # did use the whitelist here.
        if parts[-1].upper() not in ("US", "USA"):
            return True
    elif len(parts) == 2:
        # 2-part "City, XX" where XX isn't a real US state code at all.
        last = parts[-1]
        if re.fullmatch(r"[A-Za-z]{2,3}", last) and last.upper() not in US_STATE_CODES:
            return True
    return False


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
                role_type = j.get("roleType") or ""
                location = j.get("location", "")
                if (key in DEFENSE_COMPANY_BLOCKLIST
                        or DOMAIN_EXCLUDE_RE.search(one_liner)
                        or role_type in YC_EXCLUDED_ROLE_TYPES
                        or _yc_location_is_non_us(location)):
                    continue
                role_text = YC_ROLE_TYPE_TEXT.get(role_type, "") or category_text
                jobs.append({
                    "id": f"yc:{j['id']}",
                    "title": j.get("title", ""),
                    "companyName": j.get("companyName", ""),
                    "isStartup": True,
                    "website": known_websites.get(key) or (
                        f"https://www.workatastartup.com/companies/{j['companySlug']}"
                        if j.get("companySlug") else ""
                    ),
                    "location": location,
                    "descriptionText": f"{role_text} {j.get('companyOneLiner', '')}",
                    "link": j.get("applyUrl", ""),
                    "postedAt": "",
                })
        except Exception as e:
            print(f"{'YC ' + url.rsplit('/', 1)[-1]:28s} ({'yc':10s}): FAILED — {e}")
        time.sleep(0.5)
    return jobs


# ── UW Career Center's public "featured jobs" board ─────────────────────────
# careers.uw.edu mirrors a small, rotating subset of Handshake postings onto
# a public WordPress page (no login needed to view — only "Apply" redirects
# to Handshake). robots.txt allows it. This is NOT a substitute for full
# Handshake access (~9 jobs at a time, refreshed periodically, not UW's full
# catalog) — it's a small supplementary source, valuable mainly because it
# surfaces small/local WA employers that don't appear in any other source.
#
# UT Austin (12twenty@Texas) and the UC system (Handshake) were checked and
# don't have an equivalent public mirror — confirmed by visiting their
# actual career-services pages, not just inferred.
UW_JOBS_URL = (
    "https://careers.uw.edu/jobs/"
    "?ctag%5B0%5D=full-time-jobs&ctag%5B1%5D=internships"
    "&ctag%5B2%5D=part-time-jobs&ctag%5B3%5D=remote-jobs"
    "&stag[]=tech-data-gaming"
)

UW_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}


def fetch_uw() -> list[dict]:
    from bs4 import BeautifulSoup

    r = requests.get(UW_JOBS_URL, headers=UW_HEADERS, timeout=20)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    jobs = []
    for post in soup.select("#featured-jobs-list > div[id^='post-']"):
        title_link = post.select_one("h3.entry-title a")
        if not title_link:
            continue
        company = post.select_one(".company_name")
        summary = post.select_one(".entry-summary .entry-content")
        job_type_tag = post.select_one(".job-meta .entry-meta-item span:last-child")
        is_remote = bool(job_type_tag and "remote" in job_type_tag.get_text(strip=True).lower())

        jobs.append({
            "id": f"uw:{post['id']}",
            "title": title_link.get_text(strip=True),
            "companyName": company.get_text(strip=True) if company else "",
            "isStartup": False,  # unknown per-posting; not guessed for this source
            "website": "",
            "location": "Remote" if is_remote else "Washington",
            "descriptionText": summary.get_text(strip=True) if summary else "",
            "link": title_link["href"],
            "postedAt": "",
        })
    return jobs


# ── Named-company direct sources: Starbucks + American Express ────────────
# Giri asked specifically to be notified when these two post. Neither uses
# Greenhouse/Lever/Ashby, so they can't go in COMPANIES/FETCHERS as-is - and
# titles like "Software Engineer I" don't self-label "new grad"/"intern" the
# way the COMPANIES loop's INTERNSHIP_TITLE_RE requirement expects, so
# these are integrated the same way as YC/UW (no title-regex requirement at
# fetch time; SENIORITY_AVOID_RE/GRAD_DEGREE_AVOID_RE/domain/recency/
# underclassman all still apply in score_job() same as everywhere else).
#
# Both use a 2-tier API: a cheap search/list endpoint (title + location +
# date, no description) and a details endpoint (real description) called
# per candidate - it's what a details call costs that _cheap_title_survives()
# exists to avoid paying for on obviously-disqualified titles.
BIG_COMPANY_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
}


def _get_with_retry(url: str, **kwargs) -> requests.Response:
    """GET with retry-after-backoff on 429/5xx. The per-job details calls on
    these two sources are frequent enough in a burst (dozens back to back,
    ~0.2-0.5s apart) to trip basic rate limiting - seen directly against
    Starbucks' position_details endpoint during testing, even with one retry."""
    for delay in (3, 6, None):
        r = requests.get(url, timeout=20, **kwargs)
        if (r.status_code == 429 or r.status_code >= 500) and delay is not None:
            time.sleep(delay)
            continue
        r.raise_for_status()
        return r


def _cheap_title_survives(title: str) -> bool:
    """Pre-filter for 2-tier sources, applied before paying for a per-job
    details call. Not a substitute for score_job()'s real hard filters
    (which still run once description text is available) - just avoids
    fetching details for titles that would obviously fail anyway."""
    if SENIORITY_AVOID_RE.search(title) and not INTERNSHIP_TITLE_RE.search(title):
        return False
    if GRAD_DEGREE_AVOID_RE.search(title):
        return False
    if NON_TECHNICAL_ROLE_RE.search(title):
        return False
    return True


# Starbucks runs on Phenom People's "pcsx" career-site platform.
# apply.starbucks.com/robots.txt was checked directly and explicitly
# ALLOWS /api/pcsx (unlike Skillsire's disallow on /api/ - no conflict here).
STARBUCKS_COMPANY = {
    "name": "Starbucks", "token": "starbucks.com",
    "startup": False, "website": "https://www.starbucks.com",
}
STARBUCKS_QUERIES = [
    "Software Engineer", "Software Development Engineer",
    "Data Engineer", "Machine Learning Engineer",
]


def fetch_starbucks(company: dict) -> list[dict]:
    domain = company["token"]
    jobs, seen_ids = [], set()

    for query in STARBUCKS_QUERIES:
        start = 0
        while True:
            try:
                r = _get_with_retry(
                    "https://apply.starbucks.com/api/pcsx/search",
                    params={"domain": domain, "query": query, "start": start},
                    headers=BIG_COMPANY_HEADERS,
                )
                positions = r.json().get("data", {}).get("positions", [])
            except Exception as e:
                print(f"Starbucks search {query!r} @{start} FAILED: {e}")
                break
            if not positions:
                break
            for p in positions:
                pid, title = p.get("id"), p.get("name", "")
                std_locs = p.get("standardizedLocations") or []
                is_us = any(loc == "US" or loc.endswith(", US") for loc in std_locs)
                # Starbucks' search is fuzzy/broad (a huge, mostly-retail job
                # index) — a "Software Engineer" query surfaced a French-
                # language Canadian store-manager posting during testing.
                # Neither the English-only seniority regex nor
                # NON_US_LOCATION_RE (missing many Canadian provinces) would
                # have caught that reliably, so check the real structured
                # location data directly instead of trusting title regex.
                if not pid or pid in seen_ids or not _cheap_title_survives(title) or not is_us:
                    continue
                seen_ids.add(pid)
                jobs.append({
                    "id": f"starbucks:{pid}",
                    "title": title,
                    "companyName": company["name"],
                    "isStartup": company["startup"],
                    "website": company["website"],
                    "location": "; ".join(p.get("standardizedLocations") or p.get("locations") or []),
                    "descriptionText": "",
                    "link": f"https://apply.starbucks.com{p['positionUrl']}" if p.get("positionUrl") else "",
                    "postedAt": (
                        datetime.fromtimestamp(p["postedTs"], tz=timezone.utc).isoformat()
                        if p.get("postedTs") else ""
                    ),
                    "_detail_id": pid,
                })
            start += len(positions)
            if start > 200:  # safety cap - Starbucks' SWE-flavored postings are a small subset of its total job count
                break
            time.sleep(0.2)

    for j in jobs:
        pid = j.pop("_detail_id")
        try:
            r = _get_with_retry(
                "https://apply.starbucks.com/api/pcsx/position_details",
                params={"position_id": pid, "domain": domain, "hl": "en"},
                headers=BIG_COMPANY_HEADERS,
            )
            j["descriptionText"] = strip_html(r.json().get("data", {}).get("jobDescription") or "")
        except Exception as e:
            print(f"Starbucks details {pid} FAILED: {e}")
        time.sleep(0.4)
    return jobs


# Oracle Fusion Recruiting Cloud - American Express's platform, also used by
# many other large enterprises. Neither careers.americanexpress.com nor the
# underlying Oracle Cloud host (egug.fa.us2.oraclecloud.com) publish a
# robots.txt at all (checked directly - both 404), so there's no explicit
# disallow to weigh, unlike Skillsire. Config carries base_url + site_number
# per company (both visible in a careers page's HTML as data-apibaseurl /
# data-sitenumber) and careers_base for building real apply links, so more
# Oracle-ORC companies can be added later without new fetch logic.
ORACLE_ORC_QUERIES = ["Software Engineer", "Software Development Engineer", "Data Engineer", "Data Scientist"]

ORACLE_ORC_COMPANIES = [
    {
        "name": "American Express", "token": "amex",
        "base_url": "https://egug.fa.us2.oraclecloud.com:443",
        "site_number": "CX_1",
        "careers_base": "https://careers.americanexpress.com",
        "startup": False, "website": "https://www.americanexpress.com",
    },
]


def _oracle_date_to_iso(date_str: str) -> str:
    # The list endpoint gives a date-only string ("2026-08-10"), which
    # datetime.fromisoformat() parses as naive (no tzinfo) - score_job()
    # then compares it against an aware datetime.now(timezone.utc) and
    # raises TypeError. Force UTC midnight so it round-trips safely.
    if not date_str:
        return ""
    return date_str if "T" in date_str else f"{date_str}T00:00:00+00:00"


def fetch_oracle_orc(company: dict) -> list[dict]:
    base_url, site_number = company["base_url"], company["site_number"]
    jobs, seen_ids = [], set()

    for query in ORACLE_ORC_QUERIES:
        offset = 0
        while True:
            finder = (
                f"findReqs;siteNumber={site_number},"
                f"facetsList=LOCATIONS;WORK_LOCATIONS;WORKPLACE_TYPES;TITLES;CATEGORIES;ORGANIZATIONS;POSTING_DATES;FLEX_FIELDS,"
                f"limit=25,offset={offset},sortBy=POSTING_DATES_DESC,keyword={query}"
            )
            try:
                r = _get_with_retry(
                    f"{base_url}/hcmRestApi/resources/latest/recruitingCEJobRequisitions",
                    params={
                        "onlyData": "true", "finder": finder,
                        # Without `expand`, the response omits requisitionList
                        # entirely (only facet aggregations come back) - found
                        # by testing, not documented anywhere obvious.
                        "expand": "requisitionList.secondaryLocations,requisitionList.workLocation",
                    },
                    headers=BIG_COMPANY_HEADERS,
                )
                reqs = r.json()["items"][0].get("requisitionList", [])
            except Exception as e:
                print(f"{company['name']} search {query!r} @{offset} FAILED: {e}")
                break
            if not reqs:
                break
            for req in reqs:
                rid, title = req.get("Id"), req.get("Title", "")
                country = req.get("PrimaryLocationCountry") or ""
                if not rid or rid in seen_ids or not _cheap_title_survives(title):
                    continue
                if country and country != "US":
                    continue  # cheap pre-filter before paying for a details call
                seen_ids.add(rid)
                jobs.append({
                    "id": f"oracle_orc:{company['token']}:{rid}",
                    "title": title,
                    "companyName": company["name"],
                    "isStartup": company["startup"],
                    "website": company["website"],
                    "location": req.get("PrimaryLocation", "") or "",
                    "descriptionText": "",
                    "link": f"{company['careers_base']}/en/sites/{site_number}/job/{rid}",
                    "postedAt": _oracle_date_to_iso(req.get("PostedDate")),
                    "_detail_id": rid,
                })
            offset += len(reqs)
            if offset > 200:  # safety cap
                break
            time.sleep(0.2)

    for j in jobs:
        rid = j.pop("_detail_id")
        try:
            r = _get_with_retry(
                f"{base_url}/hcmRestApi/resources/latest/recruitingCEJobRequisitionDetails",
                params={"finder": f'ById;Id="{rid}",siteNumber={site_number}', "onlyData": "true"},
                headers=BIG_COMPANY_HEADERS,
            )
            detail = r.json()["items"][0]
            desc = " ".join(filter(None, [
                detail.get("ExternalDescriptionStr"),
                detail.get("ExternalQualificationsStr"),
                detail.get("ExternalResponsibilitiesStr"),
            ]))
            j["descriptionText"] = strip_html(desc)
        except Exception as e:
            print(f"{company['name']} details {rid} FAILED: {e}")
        time.sleep(0.4)
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
        seen_pairs.update((j["companyName"].lower(), j["title"].lower()) for j in qualifying_jobs)
    except Exception as e:
        print(f"{'Y Combinator (WaaS)':28s} ({'yc':10s}): FAILED — {e}")

    # Same reasoning as YC — small/local employers on UW's board rarely
    # label roles "new grad" explicitly.
    try:
        uw_jobs = fetch_uw()
        qualifying_jobs = [
            j for j in uw_jobs
            if not NON_TECHNICAL_ROLE_RE.search(j["title"])
            and (j["companyName"].lower(), j["title"].lower()) not in seen_pairs
        ]
        print(f"{'UW Career Center':28s} ({'uw':10s}): "
              f"{len(uw_jobs):4d} total, {len(qualifying_jobs):3d} undergrad-eligible")
        all_jobs.extend(qualifying_jobs)
        seen_pairs.update((j["companyName"].lower(), j["title"].lower()) for j in qualifying_jobs)
    except Exception as e:
        print(f"{'UW Career Center':28s} ({'uw':10s}): FAILED — {e}")

    # Starbucks + American Express — same reasoning as YC/UW (title doesn't
    # self-label "new grad"/"intern"); _cheap_title_survives() already ran
    # inside the fetcher before any details call, so this is just the
    # cross-source dedup + non-technical-role check every other block does.
    try:
        starbucks_jobs = fetch_starbucks(STARBUCKS_COMPANY)
        qualifying_jobs = [
            j for j in starbucks_jobs
            if not NON_TECHNICAL_ROLE_RE.search(j["title"])
            and (j["companyName"].lower(), j["title"].lower()) not in seen_pairs
        ]
        print(f"{'Starbucks':28s} ({'starbucks':10s}): "
              f"{len(starbucks_jobs):4d} total, {len(qualifying_jobs):3d} undergrad-eligible")
        all_jobs.extend(qualifying_jobs)
        seen_pairs.update((j["companyName"].lower(), j["title"].lower()) for j in qualifying_jobs)
    except Exception as e:
        print(f"{'Starbucks':28s} ({'starbucks':10s}): FAILED — {e}")

    for oracle_company in ORACLE_ORC_COMPANIES:
        try:
            oracle_jobs = fetch_oracle_orc(oracle_company)
            qualifying_jobs = [
                j for j in oracle_jobs
                if not NON_TECHNICAL_ROLE_RE.search(j["title"])
                and (j["companyName"].lower(), j["title"].lower()) not in seen_pairs
            ]
            print(f"{oracle_company['name']:28s} ({'oracle_orc':10s}): "
                  f"{len(oracle_jobs):4d} total, {len(qualifying_jobs):3d} undergrad-eligible")
            all_jobs.extend(qualifying_jobs)
            seen_pairs.update((j["companyName"].lower(), j["title"].lower()) for j in qualifying_jobs)
        except Exception as e:
            print(f"{oracle_company['name']:28s} ({'oracle_orc':10s}): FAILED — {e}")

    return all_jobs


# ── Scoring ────────────────────────────────────────────────────────────────
def score_job(job: dict) -> tuple[int, list[str]]:
    """
    Returns (score 0-100, list of matched reasons).
    Breakdown:
      - Skill keyword match   : 0-35 pts
      - Title / role fit      : 0-25 pts
      - Location preference   : 0-20 pts
      - Recency                : 0-20 pts (postings older than
        MAX_POSTING_AGE_DAYS are hard-filtered before scoring, not just
        down-scored)
      - Negative signals       : subtracted, floor 0
    """
    title = (job.get("title") or "").lower()
    location = (job.get("location") or "").lower()
    description = (job.get("descriptionText") or "").lower()
    full_text = f"{title} {description}"
    is_uw = job["id"].startswith("uw:")

    # ── Hard title filter — drop regardless of skill/keyword overlap ──────
    # Seniority: an explicit intern/co-op/fellowship/new-grad label overrides
    # this. "manager"/"lead"/"staff" etc. are meant to catch seniority, but
    # "manager" is also just the literal job title for Product roles at any
    # level — "Product Manager Summer 2027 Intern" was getting killed by the
    # bare "manager" match despite plainly self-labeling as an internship.
    seniority_hit = SENIORITY_AVOID_RE.search(title)
    if seniority_hit and not INTERNSHIP_TITLE_RE.search(title):
        return 0, [f"filtered: '{seniority_hit.group(0)}' in title"]

    # Grad-degree: never overridable — a posting can be both an internship
    # AND require a graduate degree ("... Internship ... PhD Student ...").
    grad_hit = GRAD_DEGREE_AVOID_RE.search(title)
    if grad_hit:
        return 0, [f"filtered: '{grad_hit.group(0)}' in title"]

    # ── Domain exclusion — cybersecurity / aerospace, title or description ─
    domain_hit = DOMAIN_EXCLUDE_RE.search(full_text)
    if domain_hit:
        return 0, [f"filtered: '{domain_hit.group(0)}' (excluded domain)"]

    # ── US-only — drop anything whose location names a foreign country/city ─
    non_us_hit = NON_US_LOCATION_RE.search(location)
    if non_us_hit:
        return 0, [f"filtered: '{non_us_hit.group(0)}' (non-US location)"]

    # ── Underclassman eligibility — never overridable, same treatment as
    # GRAD_DEGREE_AVOID_RE (an internship self-labeled as open to anyone can
    # still explicitly require 2+ years of school left in the description).
    underclass_hit = targets_underclassman(full_text)
    if underclass_hit:
        return 0, [f"filtered: targets underclassmen ('{underclass_hit}')"]

    # ── Recency — hard cutoff, checked here (not just scored below) so a
    # stale posting is dropped outright regardless of how well it otherwise
    # scores. Sources with no postedAt (YC, UW) can't be proven stale, so
    # days_old stays None for them and this doesn't drop anything — see
    # MAX_POSTING_AGE_DAYS' docstring.
    posted_at = job.get("postedAt")
    days_old = None
    if posted_at:
        try:
            posted_dt = datetime.fromisoformat(posted_at.replace("Z", "+00:00"))
            days_old = (datetime.now(timezone.utc) - posted_dt).days
        except Exception:
            days_old = None
    if days_old is not None and days_old > MAX_POSTING_AGE_DAYS:
        return 0, [f"filtered: posted {days_old}d ago (> {MAX_POSTING_AGE_DAYS}d limit)"]

    reasons = []
    score = 0

    # ── Skill match (0-35) ────────────────────────────────────────────────
    matched_skills = [s for s, pat in SKILL_PATTERNS if pat.search(full_text)]
    if len(matched_skills) < 2 and not is_uw:
        # A single keyword hit is too weak a signal on its own (e.g. an
        # accounting internship mentioning "Excel", or boilerplate template
        # text a company reuses across unrelated fellowship tracks) —
        # require at least 2 real overlaps before treating it as a fit.
        # UW is exempted: its one-sentence descriptions rarely hit 2 real
        # keyword matches at all, so this gate would zero out nearly every
        # UW posting regardless of genuine relevance. UW postings get a
        # pass here and are always written to the sheet (see
        # MIN_SCORE_TO_INCLUDE handling in _append_row) so Giri can see the
        # real score and judge fit himself rather than seeing nothing.
        return 0, ["filtered: insufficient technical skill overlap"]
    if matched_skills:
        skill_score = min(35, len(matched_skills) * 3)
        score += skill_score
        reasons.append(f"Skills: {', '.join(matched_skills[:6])}")

    # ── Title / role fit (0-25) ───────────────────────────────────────────
    # New Grad now scores equal to Internship (was 12 vs 25) — Giri's a
    # graduating senior targeting entry-level full-time roles, not chasing
    # internships as the primary goal anymore.
    if re.search(r"\bintern(ship)?s?\b", title):
        score += 25
        reasons.append("Internship title")
    elif re.search(r"new grad|university grad|early career", title):
        score += 25
        reasons.append("New-grad title")
    elif re.search(r"fellow(ship)?s?", title):
        score += 23
        reasons.append("Fellowship title")
    elif re.search(r"co-?op", title):
        score += 20
        reasons.append("Co-op title")
    elif job["id"].startswith("yc:"):
        # YC full-time postings rarely say "new grad" explicitly even when
        # they're accessible — SENIORITY_AVOID_RE already screened out
        # senior/staff/founding-engineer/CTO/lead titles, so what's left
        # gets partial credit rather than zero. Same 55-point bar as every
        # other source applies on top of this — no separate threshold.
        score += 23
        reasons.append("YC full-time (undergrad-eligible)")

    # ── Location (0-20) ───────────────────────────────────────────────────
    # "remote" matching as a bare substring means "Remote (IN)"/"Remote (UK)"
    # would score as if they were US-remote — a real bug found via a YC
    # posting (Gurugram, India) scoring 60% partly off this. Guard against
    # "remote (XX)" with a non-US code before granting the bonus; not
    # YC-specific — this loop runs for every source's location string.
    non_us_remote = re.search(r"remote\s*\(([a-z]{2,3})\)", location)
    for loc in PROFILE["preferred_locations"]:
        if loc in location:
            if loc == "remote" and non_us_remote and non_us_remote.group(1) not in ("us", "usa"):
                continue
            score += 20
            reasons.append(f"Location: {job.get('location')}")
            break

    # ── Recency (0-20) ────────────────────────────────────────────────────
    # Everything reaching this point already survived the MAX_POSTING_AGE_DAYS
    # hard cutoff above, so this just rewards "very fresh" over "within the
    # week" rather than gating on age at all.
    if days_old is not None:
        if days_old <= 2:
            score += 20
            reasons.append("Posted <=2d ago")
        else:
            score += 14
            reasons.append(f"Posted {days_old}d ago")
    else:
        score += 8  # unknown age (YC/UW) — neutral, can't verify freshness

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
    # UW gets a pass on the score floor too (small, valuable, low-signal
    # source by nature — see score_job's is_uw handling) — always written
    # so Giri can see the real percentage and judge fit himself, as long as
    # it wasn't hard-excluded outright (seniority/grad-degree/domain/US-only
    # all still return 0 with a "filtered: ..." reason, which stays dropped).
    is_uw = job["id"].startswith("uw:")
    if score < MIN_SCORE_TO_INCLUDE and not (is_uw and not reasons[0].startswith("filtered:")):
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
