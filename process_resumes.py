#!/usr/bin/env python3
"""
Process resumes from the `candidates-resume/` folder and append new candidates
to `Candidates List.xlsx`.

- Safe to re-run anytime: resumes already recorded (tracked by filename in the
  hidden "Source File" column, and backed up by an Email dedupe check) are
  skipped automatically.
- Extracts: Name, Total Exp, Cloud Exp., Tech, Current CTC, Company, Mobile,
  Email, Linked-in.
- Email / Mobile / LinkedIn are extracted deterministically (regex + PDF
  hyperlink annotations).
- Name / Total Exp / Cloud Exp. / Tech / Current CTC / Company are extracted
  with a local heuristic parser by default. If an OPENAI_API_KEY environment
  variable is set, an LLM call is used instead for these fields since they
  require real reasoning about employment dates/roles (more accurate than
  regex alone).

Usage:
    source .venv/bin/activate
    python process_resumes.py
"""

import os
import re
import sys
import json
from datetime import datetime
from pathlib import Path

import pdfplumber
import docx
import openpyxl

BASE_DIR = Path(__file__).resolve().parent
RESUME_DIR = BASE_DIR / "candidates-resume"
SHEET_PATH = BASE_DIR / "Candidates List.xlsx"
SHEET_NAME = "Sheet1"

HEADERS = [
    "Name", "Total Exp", "Cloud Exp.", "Tech", "Current CTC", "Company",
    "College", "Mobile", "Email", "Linked-in", "Source File", "Processed At",
]

CLOUD_KEYWORDS = [
    "aws", "amazon web services", "gcp", "google cloud", "azure",
    "microsoft azure", "cloud run", "ec2", "s3", "lambda", "kubernetes",
    "k8s", "docker", "terraform", "cloudformation", "devops", "ci/cd",
    "openshift", "firebase", "digitalocean", "oci", "oracle cloud",
]

TECH_KEYWORDS = [
    "python", "java", "javascript", "typescript", "c++", "c#", "go", "golang",
    "rust", "ruby", "php", "swift", "kotlin", "scala", "r ", "matlab",
    "react", "react.js", "angular", "vue", "node", "node.js", "django",
    "flask", "fastapi", "spring", "spring boot", ".net", "express",
    "aws", "gcp", "azure", "kubernetes", "docker", "terraform",
    "sql", "mysql", "postgresql", "postgres", "mongodb", "redis",
    "kafka", "rabbitmq", "elasticsearch", "spark", "hadoop", "airflow",
    "tensorflow", "pytorch", "scikit-learn", "pandas", "numpy",
    "html", "css", "graphql", "rest api", "microservices", "linux",
    "git", "jenkins", "ansible", "flutter", "android", "ios", "selenium",
]

EMAIL_RE = re.compile(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9.-]+")
PHONE_RE = re.compile(r"(?:\+?\d{1,3}[-.\s]?)?\d{10}")
DATE_RANGE_RE = re.compile(
    r"([A-Za-z]{3,9}\.?\s*\d{4})\s*[-–to]+\s*(Present|Current|[A-Za-z]{3,9}\.?\s*\d{4})",
    re.IGNORECASE,
)
MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7,
    "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


# --------------------------------------------------------------------------
# File reading
# --------------------------------------------------------------------------

def extract_pdf(path: Path):
    text_parts = []
    hyperlinks = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text_parts.append(page.extract_text() or "")
            try:
                hyperlinks.extend(getattr(page, "hyperlinks", []) or [])
            except Exception:
                pass  # Some malformed PDF link annotations fail to decode.
    return "\n".join(text_parts), hyperlinks


def extract_docx(path: Path):
    document = docx.Document(str(path))
    text_parts = [p.text for p in document.paragraphs]
    hyperlinks = []
    try:
        rels = document.part.rels
        for rel in rels.values():
            if "hyperlink" in rel.reltype:
                hyperlinks.append({"uri": rel.target_ref})
    except Exception:
        pass
    return "\n".join(text_parts), hyperlinks


def read_resume(path: Path):
    if path.suffix.lower() == ".pdf":
        return extract_pdf(path)
    if path.suffix.lower() in (".docx", ".doc"):
        return extract_docx(path)
    raise ValueError(f"Unsupported file type: {path.suffix}")


# --------------------------------------------------------------------------
# Deterministic extraction (email / mobile / linkedin)
# --------------------------------------------------------------------------

def find_email(text, hyperlinks):
    for link in hyperlinks:
        uri = link.get("uri", "") or ""
        if uri.lower().startswith("mailto:"):
            return uri.split(":", 1)[1].strip()
    match = EMAIL_RE.search(text)
    return match.group(0) if match else ""


def find_mobile(text):
    for match in PHONE_RE.finditer(text):
        digits = re.sub(r"\D", "", match.group(0))
        if 10 <= len(digits) <= 13:
            return match.group(0).strip()
    return ""


def find_linkedin(text, hyperlinks):
    for link in hyperlinks:
        uri = link.get("uri", "") or ""
        if "linkedin.com" in uri.lower():
            return uri
    match = re.search(r"(https?://)?(www\.)?linkedin\.com/\S+", text, re.IGNORECASE)
    if match:
        return match.group(0).rstrip(").,")
    return ""


# --------------------------------------------------------------------------
# Heuristic fallback extraction (no AI key available)
# --------------------------------------------------------------------------

def guess_name(text):
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if EMAIL_RE.search(line) or "linkedin" in line.lower() or "github" in line.lower():
            continue
        if len(line.split()) <= 5 and re.match(r"^[A-Za-z][A-Za-z.\s'-]+$", line):
            return line.title() if line.isupper() else line
        break
    return ""


def parse_month_year(token):
    token = token.strip().lower().rstrip(".")
    if token in ("present", "current"):
        now = datetime.now()
        return now.year * 12 + now.month
    match = re.match(r"([a-z]{3,9})\.?\s*(\d{4})", token)
    if not match:
        return None
    month_str, year_str = match.groups()
    month = MONTHS.get(month_str[:3])
    if not month:
        return None
    return int(year_str) * 12 + month


EXPERIENCE_HEADERS = re.compile(
    r"^\s*(work\s+)?(professional\s+)?experience\s*$|^\s*employment(\s+history)?\s*$",
    re.IGNORECASE,
)
EDUCATION_HEADERS = re.compile(r"^\s*education\s*:?\s*$", re.IGNORECASE)
SECTION_HEADERS = re.compile(
    r"^\s*(education|projects?|skills|certifications?|achievements?|"
    r"publications?|extra[- ]curricular|awards?|summary|objective)\s*:?\s*$",
    re.IGNORECASE,
)
ALL_SECTION_HEADERS = re.compile(
    r"^\s*((work\s+)?(professional\s+)?experience|employment(\s+history)?|"
    r"education|projects?|skills|certifications?|achievements?|"
    r"publications?|extra[- ]curricular|awards?|summary|objective)\s*:?\s*$",
    re.IGNORECASE,
)
COLLEGE_KEYWORDS_RE = re.compile(
    r"\b(university|institute of technology|\biit\b|\bnit\b|college|"
    r"polytechnic|school of engineering|institute)\b",
    re.IGNORECASE,
)
ROLE_KEYWORDS = (
    "engineer", "developer", "intern", "sde", "scientist", "analyst",
    "manager", "consultant", "architect", "lead", "specialist", "designer",
    "researcher", "founder", "director", "volunteer",
)


def restrict_to_experience_section(lines):
    """Return only the lines belonging to the Experience/Employment section,
    if such a section header can be found; otherwise return all lines."""
    start_idx = None
    end_idx = len(lines)
    for idx, line in enumerate(lines):
        if EXPERIENCE_HEADERS.match(line.strip()):
            start_idx = idx + 1
            break
    if start_idx is None:
        return lines
    for idx in range(start_idx, len(lines)):
        if SECTION_HEADERS.match(lines[idx].strip()):
            end_idx = idx
            break
    return lines[start_idx:end_idx]


def restrict_to_education_section(lines):
    """Return only the lines belonging to the Education section, if such a
    section header can be found; otherwise return all lines."""
    start_idx = None
    end_idx = len(lines)
    for idx, line in enumerate(lines):
        if EDUCATION_HEADERS.match(line.strip()):
            start_idx = idx + 1
            break
    if start_idx is None:
        return lines
    for idx in range(start_idx, len(lines)):
        if ALL_SECTION_HEADERS.match(lines[idx].strip()):
            end_idx = idx
            break
    return lines[start_idx:end_idx]


def guess_college(text):
    lines = restrict_to_education_section(text.splitlines())
    for line in lines:
        line = line.strip()
        if line and COLLEGE_KEYWORDS_RE.search(line):
            # Strip a trailing date range if present on the same line.
            return DATE_RANGE_RE.sub("", line).strip(" -|\t•·,")
    return ""


def clean_company_candidate(line):
    """Strip a date range out of a line and judge whether what remains looks
    like a company name rather than a job title."""
    candidate = DATE_RANGE_RE.sub("", line).strip(" -|\t•·,")
    if not candidate:
        return None
    if any(kw in candidate.lower() for kw in ROLE_KEYWORDS):
        return None
    if "experience" in candidate.lower() or "years)" in candidate.lower():
        return None
    if len(candidate.split()) > 6:
        return None
    return candidate


def extract_experience_entries(text):
    """Return list of (start_months, end_months, context_snippet, company)."""
    entries = []
    all_lines = text.splitlines()
    lines = restrict_to_experience_section(all_lines)
    for idx, line in enumerate(lines):
        m = DATE_RANGE_RE.search(line)
        if not m:
            continue
        start = parse_month_year(m.group(1))
        end = parse_month_year(m.group(2))
        if start is None or end is None or end < start:
            continue
        # Grab a bit of surrounding context (the line itself + next couple lines)
        context = " ".join(lines[max(0, idx - 1): idx + 3])

        # The company name is usually either on the same line as the date
        # (with the role on the next line) or on the line just above it.
        company = clean_company_candidate(line)
        if not company and idx > 0:
            company = clean_company_candidate(lines[idx - 1])
        entries.append((start, end, context, company or ""))
    return entries


def heuristic_extract(text):
    entries = extract_experience_entries(text)

    total_months = sum(end - start for start, end, _, _ in entries)
    cloud_months = sum(
        (end - start) for start, end, context, _ in entries
        if any(kw in context.lower() for kw in CLOUD_KEYWORDS)
    )

    lower_text = text.lower()
    tech_found = [kw for kw in TECH_KEYWORDS if kw.strip() in lower_text]
    # De-dup overlapping keywords (e.g. "node" vs "node.js")
    tech_found_sorted = sorted(set(tech_found), key=len, reverse=True)
    deduped = []
    for kw in tech_found_sorted:
        if not any(kw != other and kw in other for other in deduped):
            deduped.append(kw)
    tech = ", ".join(sorted(deduped, key=str.lower))

    # First experience entry is usually the most recent/current one.
    company = next((e[3] for e in entries if e[3]), "")

    ctc_match = re.search(r"CTC[:\s]*([\d.,]+\s*(LPA|lakh|lacs)?)", text, re.IGNORECASE)
    current_ctc = ctc_match.group(1).strip() if ctc_match else ""

    return {
        "name": guess_name(text),
        "total_exp": round(total_months / 12, 1) if total_months else None,
        "cloud_exp": round(cloud_months / 12, 1) if cloud_months else None,
        "tech": tech,
        "current_ctc": current_ctc,
        "company": company,
        "college": guess_college(text),
    }


# --------------------------------------------------------------------------
# Optional AI-assisted extraction
# --------------------------------------------------------------------------

def ai_extract(text):
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        return None
    try:
        from openai import OpenAI
    except ImportError:
        return None

    client = OpenAI(api_key=api_key)
    prompt = (
        "Extract the following fields from this resume text as strict JSON "
        "with keys: name, total_exp (number, total years of professional "
        "tech experience, estimate from employment history, null if unknown), "
        "cloud_exp (number, years of experience specifically with cloud "
        "platforms like AWS/GCP/Azure, 0 if none mentioned), "
        "tech (comma-separated list of technical skills/stack), "
        "current_ctc (string, current salary/CTC if mentioned, else empty), "
        "company (string, most recent/current employer, empty if none, "
        "e.g. for students/freshers), "
        "college (string, name of the university/college/institute from the "
        "Education section, empty if none).\n\n"
        f"Resume text:\n{text[:8000]}"
    )
    try:
        response = client.chat.completions.create(
            model=os.environ.get("OPENAI_MODEL", "gpt-4o-mini"),
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0,
        )
        data = json.loads(response.choices[0].message.content)
        return data
    except Exception as exc:
        print(f"  [warn] AI extraction failed, falling back to heuristics: {exc}")
        return None


# --------------------------------------------------------------------------
# Spreadsheet handling
# --------------------------------------------------------------------------

def load_sheet():
    if not SHEET_PATH.exists():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        ws.append(HEADERS)
        return wb, ws

    wb = openpyxl.load_workbook(SHEET_PATH)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    header_row = [cell.value for cell in ws[1]]
    for col_idx, col_name in enumerate(HEADERS, start=1):
        existing = header_row[col_idx - 1] if col_idx <= len(header_row) else None
        if existing == col_name:
            continue
        if col_name not in header_row:
            # New column expected by the script (e.g. "College") that the
            # sheet doesn't have yet: insert it at the right position so
            # existing row data doesn't get misaligned/overwritten.
            ws.insert_cols(col_idx)
            ws.cell(row=1, column=col_idx, value=col_name)
            header_row.insert(col_idx - 1, col_name)
        else:
            ws.cell(row=1, column=col_idx, value=col_name)
    return wb, ws


def get_known_keys(ws):
    known_files = set()
    known_emails = set()
    known_names_without_email = set()
    header = [cell.value for cell in ws[1]]
    try:
        name_col = header.index("Name") + 1
        email_col = header.index("Email") + 1
        file_col = header.index("Source File") + 1
    except ValueError:
        name_col, email_col, file_col = 1, 8, 10

    for row in ws.iter_rows(min_row=2):
        name_cell = row[name_col - 1].value if len(row) >= name_col else None
        email_cell = row[email_col - 1].value if len(row) >= email_col else None
        file_cell = row[file_col - 1].value if len(row) >= file_col else None
        if email_cell:
            known_emails.add(str(email_cell).strip().lower())
        elif name_cell:
            # No email on this row (likely a manually-entered candidate) -
            # fall back to matching on name to avoid re-adding them.
            known_names_without_email.add(str(name_cell).strip().lower())
        if file_cell:
            known_files.add(str(file_cell).strip())
    return known_files, known_emails, known_names_without_email


def next_empty_row(ws):
    row = 2
    while ws.cell(row=row, column=1).value is not None:
        row += 1
    return row


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def process_file(path: Path):
    text, hyperlinks = read_resume(path)
    if not text.strip():
        print(f"  [warn] No text extracted from {path.name}, skipping.")
        return None

    email = find_email(text, hyperlinks)
    mobile = find_mobile(text)
    linkedin = find_linkedin(text, hyperlinks)

    ai_data = ai_extract(text)
    if ai_data:
        name = ai_data.get("name") or guess_name(text)
        total_exp = ai_data.get("total_exp")
        cloud_exp = ai_data.get("cloud_exp")
        tech = ai_data.get("tech") or ""
        current_ctc = ai_data.get("current_ctc") or ""
        company = ai_data.get("company") or ""
        college = ai_data.get("college") or guess_college(text)
    else:
        h = heuristic_extract(text)
        name = h["name"]
        total_exp = h["total_exp"]
        cloud_exp = h["cloud_exp"]
        tech = h["tech"]
        current_ctc = h["current_ctc"]
        company = h["company"]
        college = h["college"]

    return {
        "Name": name,
        "Total Exp": total_exp,
        "Cloud Exp.": cloud_exp,
        "Tech": tech,
        "Current CTC": current_ctc,
        "Company": company,
        "College": college,
        "Mobile": mobile,
        "Email": email,
        "Linked-in": linkedin,
        "Source File": path.name,
        "Processed At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def main():
    if not RESUME_DIR.exists():
        print(f"Resume folder not found: {RESUME_DIR}")
        sys.exit(1)

    wb, ws = load_sheet()
    known_files, known_emails, known_names_without_email = get_known_keys(ws)

    resume_files = sorted(
        p for p in RESUME_DIR.iterdir()
        if p.suffix.lower() in (".pdf", ".docx", ".doc") and not p.name.startswith("~$")
    )

    new_count = 0
    row_idx = next_empty_row(ws)

    for path in resume_files:
        if path.name in known_files:
            continue

        print(f"Processing {path.name} ...")
        try:
            record = process_file(path)
        except Exception as exc:
            print(f"  [error] Failed to process {path.name}: {exc}")
            continue

        if record is None:
            continue

        email_key = record["Email"].strip().lower() if record["Email"] else ""
        name_key = record["Name"].strip().lower() if record["Name"] else ""
        if email_key and email_key in known_emails:
            print(f"  [skip] Duplicate email already in sheet: {record['Email']}")
            known_files.add(path.name)
            continue
        if not email_key and name_key and name_key in known_names_without_email:
            print(f"  [skip] Candidate '{record['Name']}' already in sheet (matched by name).")
            known_files.add(path.name)
            continue

        for col_idx, col_name in enumerate(HEADERS, start=1):
            ws.cell(row=row_idx, column=col_idx, value=record.get(col_name))

        if email_key:
            known_emails.add(email_key)
        elif name_key:
            known_names_without_email.add(name_key)
        known_files.add(path.name)
        row_idx += 1
        new_count += 1

    if new_count:
        wb.save(SHEET_PATH)
        print(f"\nSaved {new_count} new candidate(s) to {SHEET_PATH.name}")
    else:
        print("\nNo new resumes to add. Everything is already up to date.")


if __name__ == "__main__":
    main()
