"""
update_resume.py
----------------
Downloads Christopher Ward's resume .docx from Google Drive,
parses the content, and injects it into the placeholder zones
in index.html.

Placeholder zones in index.html:
  <!-- SUMMARY_START -->    ... <!-- SUMMARY_END -->
  <!-- EXPERIENCE_START --> ... <!-- EXPERIENCE_END -->
  <!-- EDUCATION_START -->  ... <!-- EDUCATION_END -->
  <!-- SKILLS_START -->     ... <!-- SKILLS_END -->
  <!-- PROJECTS_START -->   ... <!-- PROJECTS_END -->

Environment variables required:
  GDRIVE_FILE_ID  — the ID portion of your Google Drive share URL

Usage:
  pip install python-docx requests
  GDRIVE_FILE_ID=your_id_here python scripts/update_resume.py
"""

import os
import re
import sys
import html
import requests
import tempfile
from docx import Document


# ── Config ────────────────────────────────────────────────────────────────────

GDRIVE_FILE_ID = os.environ.get("GDRIVE_FILE_ID")
GDRIVE_EXPORT_URL = f"https://drive.google.com/uc?export=download&id={GDRIVE_FILE_ID}"
HTML_FILE = "index.html"

# Section heading text as it appears in the Word doc (case-insensitive match)
SECTION_HEADINGS = {
    "summary":    ["summary", "professional summary"],
    "experience": ["experience", "work experience", "professional experience"],
    "education":  ["education"],
    "skills":     ["skills", "technical skills"],
    "projects":   ["projects", "projects & leadership", "projects and leadership"],
}

# Lines to ignore entirely — footer content, citizenship notes, availability, etc.
IGNORED_LINES = [
    "u.s. citizen",
    "us citizen",
    "available",
    "available to work",
]


# ── Download ──────────────────────────────────────────────────────────────────

def download_docx(file_id: str) -> str:
    """Download the .docx from Google Drive, return path to temp file."""
    if not file_id:
        print("ERROR: GDRIVE_FILE_ID environment variable is not set.")
        sys.exit(1)

    print(f"Downloading resume from Google Drive (id: {file_id})...")
    url = f"https://drive.google.com/uc?export=download&id={file_id}"

    session = requests.Session()
    response = session.get(url, stream=True, timeout=30)

    # Google Drive sometimes returns a confirmation page for large files
    for key, value in response.cookies.items():
        if key.startswith("download_warning"):
            params = {"id": file_id, "confirm": value}
            response = session.get(url, params=params, stream=True, timeout=30)
            break

    response.raise_for_status()

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    for chunk in response.iter_content(chunk_size=32768):
        if chunk:
            tmp.write(chunk)
    tmp.close()
    print(f"  Saved to {tmp.name}")
    return tmp.name


# ── Parse .docx ───────────────────────────────────────────────────────────────

def get_paragraph_text(para) -> str:
    """Return the plain text of a paragraph."""
    return para.text.strip()


def is_section_heading(text: str, section_key: str) -> bool:
    """Check if a paragraph text matches a known section heading."""
    lower = text.lower().strip(":")
    return lower in SECTION_HEADINGS[section_key]


def is_any_heading(text: str) -> bool:
    """Check if text matches any section heading (to detect section boundaries)."""
    lower = text.lower().strip(":")
    for aliases in SECTION_HEADINGS.values():
        if lower in aliases:
            return True
    return False


def extract_sections(docx_path: str) -> dict:
    """
    Walk the document paragraphs and bucket them by section.
    Returns a dict: { 'summary': [...paras], 'experience': [...paras], ... }
    """
    doc = Document(docx_path)
    sections = {k: [] for k in SECTION_HEADINGS}
    current_section = None

    for para in doc.paragraphs:
        text = get_paragraph_text(para)
        if not text:
            continue

        # Skip footer/noise lines
        if any(text.lower().startswith(ignored) for ignored in IGNORED_LINES):
            continue

        # Check if this paragraph is a section heading
        matched_section = None
        for key in SECTION_HEADINGS:
            if is_section_heading(text, key):
                matched_section = key
                break

        if matched_section:
            current_section = matched_section
            continue  # Don't add the heading itself

        if current_section:
            sections[current_section].append(para)

    return sections


# ── Build HTML blocks ─────────────────────────────────────────────────────────

def esc(text: str) -> str:
    """HTML-escape a string and normalize special characters."""
    text = text.replace("\u2013", "–").replace("\u2014", "—")
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    return html.escape(text, quote=False)


def is_bullet(para) -> bool:
    """Return True if the paragraph is a list item."""
    style = para.style.name.lower()
    return "list" in style or para.paragraph_format.left_indent is not None and \
           para.paragraph_format.left_indent > 0


def para_is_bold_label(para) -> bool:
    """Return True if the paragraph starts with a bold run (skill category line)."""
    for run in para.runs:
        if run.bold and run.text.strip():
            return True
    return False


def build_summary_html(paras: list) -> str:
    lines = [p.text.strip() for p in paras if p.text.strip()]
    if not lines:
        return '<p class="summary-text">Summary not found in document.</p>'
    return f'<p class="summary-text">{esc(" ".join(lines))}</p>\n'


def build_skills_html(paras: list) -> str:
    """
    Expects paragraphs like:
      "Networking & Infrastructure: Layer 1-3 troubleshooting, ..."
    Bold run = category name, rest = items.
    """
    html_parts = ['<div class="skills-grid">\n']

    for para in paras:
        text = para.text.strip()
        if not text:
            continue

        # Split on first colon to get category : items
        if ":" in text:
            category, _, items = text.partition(":")
            category = category.strip()
            items = items.strip()
        else:
            category = text
            items = ""

        # Escape ampersands for HTML
        category_html = esc(category).replace("&amp;", "&amp;")
        items_html = esc(items)

        html_parts.append(
            f'          <div class="skill-group">\n'
            f'            <span class="skill-category">{category_html}</span>\n'
            f'            <span class="skill-items">{items_html}</span>\n'
            f'          </div>\n'
        )

    html_parts.append('          </div>')
    return "".join(html_parts)


def build_entry_html(org: str, role: str, location: str, date: str, bullets: list, detail: str = "") -> str:
    """Build a single .entry block."""
    role_html = f'\n                <span class="entry-role">{esc(role)}</span>' if role else ""
    location_html = f'\n                <span class="entry-location">{esc(location)}</span>' if location else ""

    bullets_html = ""
    if bullets:
        items = "\n".join(f'              <li>{esc(b)}</li>' for b in bullets if b)
        bullets_html = f'\n            <ul class="entry-list">\n{items}\n            </ul>'

    detail_html = f'\n            <p class="entry-detail">{esc(detail)}</p>' if detail else ""

    return (
        f'          <div class="entry">\n'
        f'            <div class="entry-header">\n'
        f'              <div class="entry-left">\n'
        f'                <span class="entry-org">{esc(org)}</span>{role_html}\n'
        f'              </div>\n'
        f'              <div class="entry-right">{location_html}\n'
        f'                <span class="entry-date">{esc(date)}</span>\n'
        f'              </div>\n'
        f'            </div>{bullets_html}{detail_html}\n'
        f'          </div>\n'
    )


def parse_job_header(text: str):
    """
    Try to split a job header line like:
      "Planet Networks | Newton, NJ | Aug 2025 – Present"
      "Tier 2 Customer Success Associate"
    Returns (org, location, date) or None if it doesn't look like a header.
    """
    parts = [p.strip() for p in text.split("|")]
    if len(parts) >= 3:
        return parts[0], parts[1], parts[2]
    elif len(parts) == 2:
        return parts[0], "", parts[1]
    return None


def build_experience_html(paras: list) -> str:
    """
    Parses experience paragraphs into entry blocks.
    Handles the two-line format from the Word doc:
      Line 1: "Org | Location | Date"
      Line 2: "Role Title"  (bold, non-bullet)
      Lines 3+: bullet points
    """
    entries = []
    i = 0

    while i < len(paras):
        para = paras[i]
        text = para.text.strip()
        if not text:
            i += 1
            continue

        header = parse_job_header(text)
        if header:
            org, location, date = header
            role = ""
            bullets = []

            i += 1
            # Next non-empty paragraph might be the role title
            while i < len(paras) and not paras[i].text.strip():
                i += 1
            if i < len(paras):
                next_text = paras[i].text.strip()
                next_header = parse_job_header(next_text)
                if next_header is None and not is_bullet(paras[i]):
                    role = next_text
                    i += 1

            # Collect bullets until next header or end
            while i < len(paras):
                pt = paras[i].text.strip()
                if not pt:
                    i += 1
                    continue
                if parse_job_header(pt):
                    break
                if is_bullet(paras[i]) or pt.startswith("-") or pt.startswith("•"):
                    bullets.append(pt.lstrip("-•– ").strip())
                    i += 1
                else:
                    # Could be role title if we haven't set one yet
                    if not role:
                        role = pt
                        i += 1
                    else:
                        break

            entries.append(build_entry_html(org, role, location, date, bullets))
        else:
            # Paragraph doesn't look like a header — skip or treat as orphan bullet
            i += 1

    return "\n".join(entries) if entries else "<!-- No experience entries parsed -->\n"


def build_education_html(paras: list) -> str:
    """Parse education section."""
    org = role = location = date = detail = ""
    bullets = []

    for para in paras:
        text = para.text.strip()
        if not text:
            continue
        header = parse_job_header(text)
        if header:
            org, location, date = header
        elif "bachelor" in text.lower() or "b.s." in text.lower() or "degree" in text.lower():
            role = text
        elif "gpa" in text.lower() or "coursework" in text.lower():
            detail = text
        elif is_bullet(para):
            bullets.append(text.lstrip("-•– ").strip())

    if not org:
        # Fallback: use first paragraph as org
        for para in paras:
            if para.text.strip():
                org = para.text.strip()
                break

    return build_entry_html(org, role, location, date, bullets, detail)


def build_projects_html(paras: list) -> str:
    """Parse projects & leadership section — same logic as experience."""
    return build_experience_html(paras)


# ── Inject into HTML ──────────────────────────────────────────────────────────

def inject_section(html_content: str, zone: str, new_content: str) -> str:
    """Replace everything between <!-- ZONE_START --> and <!-- ZONE_END -->."""
    start_tag = f"<!-- {zone}_START -->"
    end_tag   = f"<!-- {zone}_END -->"

    pattern = re.compile(
        rf"({re.escape(start_tag)})(.*?)({re.escape(end_tag)})",
        re.DOTALL
    )

    if not pattern.search(html_content):
        print(f"  WARNING: Placeholder '{zone}' not found in HTML. Skipping.")
        return html_content

    replacement = f"{start_tag}\n{new_content}\n          {end_tag}"
    return pattern.sub(replacement, html_content)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    # 1. Download
    docx_path = download_docx(GDRIVE_FILE_ID)

    try:
        # 2. Parse
        print("Parsing document...")
        sections = extract_sections(docx_path)

        for key, paras in sections.items():
            print(f"  [{key}] — {len(paras)} paragraph(s) found")

        # 3. Build HTML blocks
        print("Building HTML content...")
        summary_html    = build_summary_html(sections["summary"])
        experience_html = build_experience_html(sections["experience"])
        education_html  = build_education_html(sections["education"])
        skills_html     = build_skills_html(sections["skills"])
        projects_html   = build_projects_html(sections["projects"])

        # 4. Load index.html
        if not os.path.exists(HTML_FILE):
            print(f"ERROR: {HTML_FILE} not found. Run this script from the repo root.")
            sys.exit(1)

        with open(HTML_FILE, "r", encoding="utf-8") as f:
            content = f.read()

        original = content

        # 5. Inject
        print("Injecting content into index.html...")
        content = inject_section(content, "SUMMARY",    summary_html)
        content = inject_section(content, "EXPERIENCE", experience_html)
        content = inject_section(content, "EDUCATION",  education_html)
        content = inject_section(content, "SKILLS",     skills_html)
        content = inject_section(content, "PROJECTS",   projects_html)

        # 6. Write back only if changed
        if content == original:
            print("No changes detected — index.html is already up to date.")
        else:
            with open(HTML_FILE, "w", encoding="utf-8") as f:
                f.write(content)
            print("index.html updated successfully.")

    finally:
        os.unlink(docx_path)


if __name__ == "__main__":
    main()
