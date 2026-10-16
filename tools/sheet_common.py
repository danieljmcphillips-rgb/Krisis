"""
Column layout shared by sheet_export.py and sheet_import.py.

The Google Sheet has two working tabs:

  Directory  one row per record in resources.json, exactly as published.
             Edit a cell here to correct an existing record.
  Add new    one row per SERVICE (not per council). List every council it
             serves in one cell, separated by semicolons, or write
             "Scotland-wide". The importer turns it into one record per council.

Headers are friendly labels; FIELD maps each label to its resources.json key.
"""
import re

NATIONAL = "Scotland-wide"

# (sheet header, resources.json key, column width)
DIRECTORY_COLS = [
    ("ID", "id", 8),
    ("Category", "category", 26),
    ("Council", "local_authority", 22),
    ("Service name", "name", 38),
    ("Description", "description", 50),
    ("Phone", "phone", 16),
    ("Email", "email", 30),
    ("Website", "url", 40),
    ("Opening times", "opening_times", 30),
    ("Who it's for / key notes", "notes", 50),
    ("More detail", "notes_detail", 50),
    ("Hours by location", "notes_list", 30),
    ("Source page", "source_url", 40),
    ("Last checked", "last_checked", 12),
    ("Publish", "publish", 9),
    ("Verified", "verified", 9),
    ("Hide contact details", "contact_hidden", 10),
]

ADD_NEW_COLS = [
    ("Category", "category", 26),
    ("Councils served", "councils", 30),
    ("Service name", "name", 38),
    ("Description", "description", 50),
    ("Phone", "phone", 16),
    ("Email", "email", 30),
    ("Website", "url", 40),
    ("Opening times", "opening_times", 30),
    ("Who it's for / key notes", "notes", 50),
    ("More detail", "notes_detail", 50),
    ("Phone type", "contact_type", 14),
    ("Age range (their words)", "age_range", 22),
    ("Source page", "source_url", 40),
    ("Date checked", "date_checked", 12),
    ("Check result", "status", 12),
    ("Added by", "added_by", 14),
]

PUBLISH = ["yes", "no", "review"]
YES_BLANK = ["yes", ""]
CONTACT_TYPES = ["helpline", "switchboard", "referral line", "none"]
STATUSES = ["ok", "link-only", "dead", "moved", "unclear"]
OFFICE_NOTE = "Office number, not a support line."

DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
PHONE_RE = re.compile(r"^(0\d[\d ]{8,12}|\d{3}( \d{3})?)$")
EMAIL_RE = re.compile(r"^[\w.+\-]+@[\w\-]+(\.[\w\-]+)+$")


def notes_list_to_text(items):
    """[{'label','text'}] -> one 'Label = text' per line."""
    return "\n".join(f"{i['label']} = {i['text']}" for i in items or [])


def text_to_notes_list(text):
    out = []
    for line in str(text or "").splitlines():
        if not line.strip():
            continue
        if " = " not in line:
            raise ValueError(f"'Hours by location' line needs 'Label = text': {line!r}")
        label, txt = line.split(" = ", 1)
        out.append({"label": label.strip(), "text": txt.strip()})
    return out


def cell_str(v):
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()
