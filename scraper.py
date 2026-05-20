"""
Giri Vignesh — Daily Job Tracker
Scrapes LinkedIn via Apify, scores jobs against your profile,
and writes/updates jobs.xlsx in the repo.
"""

import os
import json
import time
from datetime import datetime
from apify_client import ApifyClient
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

# ── Profile: Giri's skills and preferences ────────────────────────────────
PROFILE = {
    "skills": [
        "python", "javascript", "react", "flask", "node", "aws",
        "mongodb", "rest", "api", "html", "css", "docker", "git",
        "sql", "tensorflow", "pytorch", "keras", "nlp", "machine learning",
        "full stack", "full-stack", "backend", "frontend", "c++", "c",
        "next.js", "nextjs", "express", "llm", "agile", "ci/cd",
        "postman", "oop", "data structures", "algorithms", "typescript"
    ],
    "preferred_locations": ["washington", "seattle", "bellevue", "redmond",
                             "remote", "texas", "dallas", "austin", "houston"],
    "target_titles": ["software engineer", "software developer", "swe",
                      "full stack", "frontend", "backend", "junior developer",
                      "jr developer", "jr. developer", "intern"],
    "avoid_titles": ["senior", "staff", "principal", "manager", "director",
                     "phd", "research scientist", "data scientist", "devops",
                     "security engineer", "hardware"],
    "preferred_levels": ["internship", "entry level", "entry_level",
                          "associate", "junior"],
}

# ── Search URLs ────────────────────────────────────────────────────────────
SEARCH_URLS = [
    # Washington internships
    "https://www.linkedin.com/jobs/search/?keywords=software+engineer+intern&location=Washington+State&f_E=1%2C2&f_WT=1%2C2",
    # Remote internships / junior
    "https://www.linkedin.com/jobs/search/?keywords=software+engineer+intern&f_WT=2&f_E=1%2C2",
    "https://www.linkedin.com/jobs/search/?keywords=junior+software+developer&f_WT=2&f_E=1%2C2",
    # Texas
    "https://www.linkedin.com/jobs/search/?keywords=junior+software+developer&location=Texas&f_E=1%2C2",
    "https://www.linkedin.com/jobs/search/?keywords=software+engineer+intern&location=Texas&f_E=1%2C2",
    # Full stack remote
    "https://www.linkedin.com/jobs/search/?keywords=junior+full+stack+developer&f_WT=2&f_E=1%2C2",
]

# ── Scoring ────────────────────────────────────────────────────────────────
def score_job(job: dict) -> tuple[int, list[str]]:
    """
    Returns (score 0-100, list of matched keywords).
    Breakdown:
      - Skill keyword match  : 0-40 pts
      - Location preference  : 0-20 pts
      - Title match          : 0-20 pts
      - Applicant count      : 0-20 pts (fewer = better)
    """
    reasons = []
    score = 0

    title       = (job.get("title") or "").lower()
    company     = (job.get("companyName") or "").lower()
    location    = (job.get("location") or "").lower()
    description = (job.get("descriptionText") or "").lower()
    level       = (job.get("seniorityLevel") or "").lower()
    applicants  = job.get("applicantsCount") or 999
    full_text   = f"{title} {description}"

    # ── Avoid filter ──────────────────────────────────────────────────────
    for bad in PROFILE["avoid_titles"]:
        if bad in title:
            return 0, [f"filtered: '{bad}' in title"]

    # ── Skill match (0-40) ────────────────────────────────────────────────
    matched_skills = []
    for skill in PROFILE["skills"]:
        if skill in full_text:
            matched_skills.append(skill)
    skill_score = min(40, len(matched_skills) * 4)
    score += skill_score
    if matched_skills:
        reasons.append(f"Skills: {', '.join(matched_skills[:6])}")

    # ── Location (0-20) ───────────────────────────────────────────────────
    for loc in PROFILE["preferred_locations"]:
        if loc in location:
            score += 20
            reasons.append(f"Location: {job.get('location')}")
            break

    # ── Title match (0-20) ────────────────────────────────────────────────
    title_score = 0
    for t in PROFILE["target_titles"]:
        if t in title:
            title_score = 20
            reasons.append(f"Title match: {job.get('title')}")
            break
    score += title_score

    # ── Applicant count (0-20) ────────────────────────────────────────────
    if isinstance(applicants, int):
        if applicants < 30:
            app_score = 20
        elif applicants < 75:
            app_score = 16
        elif applicants < 125:
            app_score = 12
        elif applicants < 175:
            app_score = 6
        else:
            app_score = 2
        score += app_score
        reasons.append(f"{applicants} applicants")

    return min(score, 100), reasons


def chance_label(score: int) -> str:
    if score >= 65: return "High"
    if score >= 45: return "Medium"
    if score >= 25: return "Low"
    return "Very Low"


def chance_color(score: int) -> str:
    """Returns hex fill color for the chance cell."""
    if score >= 65: return "C6EFCE"   # green
    if score >= 45: return "FFEB9C"   # yellow
    if score >= 25: return "FFCC99"   # orange
    return "FFC7CE"                    # red


# ── Apify scrape ──────────────────────────────────────────────────────────
def scrape_jobs(api_key: str) -> list[dict]:
    client = ApifyClient(api_key)
    print("Starting Apify scrape...")

    run = client.actor("curious_coder/linkedin-jobs-scraper").call(
        run_input={
            "urls": SEARCH_URLS,
            "count": 25,
            "scrapeCompany": False,
        }
    )

    items = []
    for item in client.dataset(run["defaultDatasetId"]).iterate_items():
        items.append(item)

    print(f"Scraped {len(items)} jobs.")
    return items


# ── Deduplicate against existing sheet ───────────────────────────────────
def load_existing_ids(path: str) -> set:
    if not os.path.exists(path):
        return set()
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        ids = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0]:
                ids.add(str(row[0]))
        return ids
    except Exception:
        return set()


# ── Excel builder ─────────────────────────────────────────────────────────
HEADERS = [
    "Job ID", "Date Found", "Title", "Company", "Location",
    "Match %", "Chance", "Applicants", "Employment Type",
    "Key Skills Matched", "Link", "Status", "Notes"
]

COL_WIDTHS = [14, 12, 32, 22, 20, 10, 10, 12, 16, 36, 14, 14, 24]

NAVY   = "1A3A5C"
WHITE  = "FFFFFF"
LIGHT  = "F2F5F9"
BORDER_COLOR = "CCCCCC"

def thin_border():
    s = Side(style="thin", color=BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)


def write_excel(jobs_scored: list[dict], path: str):
    """
    If jobs.xlsx exists: append new rows only (dedup by job ID).
    If not: create fresh with header + all jobs.
    """
    existing_ids = load_existing_ids(path)

    if os.path.exists(path):
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        new_count = 0
        for job in jobs_scored:
            job_id = str(job.get("id", ""))
            if job_id in existing_ids:
                continue
            _append_row(ws, job)
            new_count += 1
        print(f"Added {new_count} new jobs to existing sheet.")
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Job Matches"
        _write_header(ws)
        for job in jobs_scored:
            _append_row(ws, job)
        print(f"Created new sheet with {len(jobs_scored)} jobs.")

    # freeze top row, set column widths
    ws.freeze_panes = "A2"
    for i, width in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # auto-filter
    ws.auto_filter.ref = ws.dimensions

    wb.save(path)
    print(f"Saved to {path}")


def _write_header(ws):
    for col, header in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font       = Font(bold=True, color=WHITE, name="Arial", size=10)
        cell.fill       = PatternFill("solid", fgColor=NAVY)
        cell.alignment  = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border     = thin_border()
    ws.row_dimensions[1].height = 22


def _append_row(ws, job: dict):
    score, reasons = score_job(job)
    if score == 0:
        return  # filtered out

    row = ws.max_row + 1
    fill_color = LIGHT if row % 2 == 0 else WHITE

    values = [
        str(job.get("id", "")),
        datetime.today().strftime("%Y-%m-%d"),
        job.get("title", ""),
        job.get("companyName", ""),
        job.get("location", ""),
        score,
        chance_label(score),
        job.get("applicantsCount", ""),
        job.get("employmentType", ""),
        ", ".join(reasons[:3]),
        job.get("link") or job.get("jobUrl", ""),
        "Not Applied",
        ""
    ]

    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font      = Font(name="Arial", size=9)
        cell.alignment = Alignment(vertical="center", wrap_text=(col in [3, 10]))
        cell.border    = thin_border()

        # row zebra stripe
        if col not in [6, 7]:
            cell.fill = PatternFill("solid", fgColor=fill_color)

    # Match % — colored by score
    pct_cell = ws.cell(row=row, column=6)
    pct_cell.value     = f"{score}%"
    pct_cell.font      = Font(name="Arial", size=9, bold=True)
    pct_cell.fill      = PatternFill("solid", fgColor=chance_color(score))
    pct_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Chance label
    ch_cell = ws.cell(row=row, column=7)
    ch_cell.fill      = PatternFill("solid", fgColor=chance_color(score))
    ch_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Link — make it clickable
    link = values[10]
    if link:
        link_cell = ws.cell(row=row, column=11, value="Apply →")
        link_cell.hyperlink = link
        link_cell.font      = Font(name="Arial", size=9, color="185FA5", underline="single")
        link_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[row].height = 18


# ── Main ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    api_key = os.environ.get("APIFY_API_KEY")
    if not api_key:
        raise ValueError("APIFY_API_KEY environment variable not set.")

    output_path = os.path.join(os.path.dirname(__file__), "jobs.xlsx")

    jobs = scrape_jobs(api_key)
    write_excel(jobs, output_path)

    print(f"\n✅ Done — {datetime.today().strftime('%Y-%m-%d %H:%M')} PST")
