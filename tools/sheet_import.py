#!/usr/bin/env python3
"""
Bring the Google Sheet back into resources.json.

    python3 tools/sheet_import.py resources.json downloaded.xlsx            # report only
    python3 tools/sheet_import.py resources.json downloaded.xlsx --apply    # write changes

Download the sheet from Google Sheets with File > Download > Microsoft Excel.

Directory tab: cells that differ from resources.json are listed, and applied
  with --apply. A record missing from the sheet is reported, never deleted.
Add new tab: rows with Check result ok or link-only become one record per
  council served (or one national record). Anything that fails a check is
  listed under NOT ADDED with the reason -- nothing is guessed.

Report-only is the default, so read the report before running with --apply.
"""
import argparse
import json
import os
import re
import sys
from datetime import datetime, timedelta

from openpyxl import load_workbook

from councils import canon_la
from sheet_common import (ADD_NEW_COLS, CONTACT_TYPES, DATE_RE, DIRECTORY_COLS,
                          EMAIL_RE, NATIONAL, OFFICE_NOTE, PHONE_RE, PUBLISH,
                          STATUSES, cell_str, notes_list_to_text,
                          text_to_notes_list)

HERE = os.path.dirname(os.path.abspath(__file__))
LAST_ID_FILE = os.path.join(HERE, "last_id.txt")
ORDER = ["id", "scope", "la_code", "local_authority", "category", "name",
         "description", "phone", "email", "url", "opening_times", "notes",
         "source_url", "last_checked", "publish", "verified", "contact_hidden",
         "notes_detail"]


def norm_name(s):
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def read_tab(wb, tab, cols):
    if tab not in wb.sheetnames:
        sys.exit(f"Tab {tab!r} missing from the workbook.")
    ws = wb[tab]
    got = [cell_str(c.value) for c in ws[1]][:len(cols)]
    want = [c[0] for c in cols]
    if got != want:
        sys.exit(f"{tab} headers changed.\n  expected {want}\n  found    {got}")
    rows = []
    for r in range(2, ws.max_row + 1):
        vals = {key: cell_str(ws.cell(row=r, column=i).value)
                for i, (_, key, _) in enumerate(cols, 1)}
        if any(vals.values()):
            rows.append((r, vals))
    return rows


def resolve_category(v, slugs, by_label):
    return v if v in slugs else by_label.get(v.lower())


def contact_problems(phone, email, url, source):
    p = []
    if phone and not PHONE_RE.match(phone):
        p.append(f"phone {phone!r} doesn't look like a UK number")
    if email and not EMAIL_RE.match(email):
        p.append(f"email {email!r} is malformed")
    for label, u in (("website", url), ("source page", source)):
        if u and not u.startswith(("http://", "https://")):
            p.append(f"{label} must start with https://")
    return p


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("resources")
    ap.add_argument("workbook")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--force", action="store_true",
                    help="import even though resources.json changed since the sheet was exported")
    a = ap.parse_args()

    data = json.load(open(a.resources, encoding="utf-8"))
    slugs = {c["slug"] for c in data["categories"]}
    by_label = {c["label"].lower(): c["slug"] for c in data["categories"]}
    la_code = {c["name"]: c["code"] for c in data["councils"]}
    recs = {r["id"]: r for r in data["resources"]}
    wb = load_workbook(a.workbook)
    base = wb["Lists"]["H2"].value if "Lists" in wb.sheetnames else ""
    # Google Sheets may turn the stamp into a real date on upload
    base = base.strftime("%Y-%m-%dT%H:%M:%S") if hasattr(base, "strftime") else cell_str(base)
    if base != data.get("generated", "") and not a.force:
        sys.exit(f"This sheet was built from resources.json generated {base or '(unknown)'}, but the "
                 f"current file is generated {data.get('generated')}. Someone changed resources.json "
                 "since the sheet was exported, and importing would undo their changes on the Directory "
                 "tab. Copy your Add new rows into a fresh export instead. (--force overrides this.)")

    errors, edits = [], []

    # ------------------------------------------------ Directory: edits
    seen = set()
    for rownum, v in read_tab(wb, "Directory", DIRECTORY_COLS):
        where = f"Directory row {rownum}"
        rid = v["id"]
        if not rid:
            errors.append(f"{where}: no ID. New services go on the 'Add new' tab.")
            continue
        if rid not in recs:
            errors.append(f"{where}: ID {rid} is not in resources.json. IDs must not be typed by hand.")
            continue
        if rid in seen:
            errors.append(f"{where}: ID {rid} appears twice.")
            continue
        seen.add(rid)
        r = recs[rid]
        new = dict(v)
        bad = []
        cat = resolve_category(v["category"], slugs, by_label)
        if not cat:
            bad.append(f"unknown category {v['category']!r}")
        new["category"] = cat
        if v["local_authority"] == NATIONAL:
            new.update(scope="national", local_authority="")
        else:
            la = canon_la(v["local_authority"])
            if not la:
                bad.append(f"unknown council {v['local_authority']!r}")
            new.update(scope="local", local_authority=la)
        if v["publish"] not in PUBLISH:
            bad.append(f"Publish must be one of {PUBLISH}")
        for k in ("verified", "contact_hidden"):
            if v[k] not in ("yes", "no", ""):
                bad.append(f"{k} must be yes or blank")
        if v["last_checked"] and not DATE_RE.match(v["last_checked"]):
            bad.append("Last checked must be YYYY-MM-DD")
        # only check contact values that were changed on the sheet
        bad += contact_problems(*(v[k] if v[k] != (r.get(k) or "") else ""
                                  for k in ("phone", "email", "url")), "")
        try:
            new["notes_list"] = text_to_notes_list(v["notes_list"])
        except ValueError as e:
            bad.append(str(e))
        if bad:
            errors.append(f"{where} ({rid} {r['name']}): " + "; ".join(bad))
            continue
        changes = {}
        for _, key, _ in DIRECTORY_COLS:
            if key == "id":
                continue
            old = r.get(key)
            if key == "notes_list":
                if notes_list_to_text(old) != notes_list_to_text(new[key]):
                    changes[key] = (old, new[key])
                continue
            if (old or "") != (new[key] or ""):
                changes[key] = (old, new[key])
        if (r.get("scope") or "") != new["scope"]:
            changes["scope"] = (r.get("scope"), new["scope"])
        if changes:
            if {"phone", "email", "url"} & changes.keys() and "last_checked" not in changes:
                errors.append(f"WARNING {where} ({rid} {r['name']}): contact details changed but "
                              "Last checked did not. Applied anyway; update the date if you checked it.")
            edits.append((rid, r["name"], changes))

    missing = [i for i in recs if i not in seen]

    # ------------------------------------------------ Add new
    existing_names = {norm_name(r["name"]) for r in data["resources"]}
    batch_names = set()
    to_add, not_added, info = [], [], []
    for rownum, v in read_tab(wb, "Add new", ADD_NEW_COLS):
        where = f"Add new row {rownum} ({v['name'] or 'no name'})"
        why = []
        status = v["status"].lower()
        if status not in STATUSES:
            not_added.append(f"{where}: Check result is {v['status']!r}; needs ok or link-only.")
            continue
        if status not in ("ok", "link-only"):
            not_added.append(f"{where}: Check result is {status}; nothing to add.")
            continue
        cat = resolve_category(v["category"], slugs, by_label)
        if not cat:
            why.append(f"unknown category {v['category']!r}")
        for k, label in (("name", "Service name"), ("url", "Website"),
                         ("source_url", "Source page"), ("date_checked", "Date checked")):
            if not v[k]:
                why.append(f"{label} is blank")
        if v["date_checked"] and not DATE_RE.match(v["date_checked"]):
            why.append("Date checked must be YYYY-MM-DD")
        if v["contact_type"] and v["contact_type"] not in CONTACT_TYPES:
            why.append(f"Phone type must be one of {CONTACT_TYPES}")
        if status == "ok" and v["phone"] and not v["contact_type"]:
            why.append("Phone type is blank: say whether the number is a helpline or a switchboard")
        why += contact_problems(v["phone"], v["email"], v["url"], v["source_url"])
        nm = norm_name(v["name"])
        if nm in existing_names:
            why.append("a record with this name is already in resources.json (correct it on the Directory tab instead)")
        elif nm in batch_names:
            why.append("same service appears twice on Add new; list all its councils in one row")
        councils = []
        parts = [p.strip() for p in re.split(r"[;\n]", v["councils"]) if p.strip()]
        if not parts:
            why.append("Councils served is blank")
        for p in parts:
            if p.lower() in ("scotland-wide", "scotland wide", "national", "scotland"):
                councils.append(NATIONAL)
                continue
            la = canon_la(p)
            if not la:
                why.append(f"council {p!r} not recognised")
            else:
                if la != p:
                    info.append(f"{where}: read {p!r} as {la}")
                councils.append(la)
        if NATIONAL in councils and len(set(councils)) > 1:
            why.append(f"mixes {NATIONAL} with named councils")
        if why:
            not_added.append(f"{where}: " + "; ".join(why))
            continue
        batch_names.add(nm)
        ok = status == "ok"
        notes = v["notes"]
        if ok and v["contact_type"] == "switchboard" and OFFICE_NOTE.lower() not in notes.lower():
            notes = (notes + " " + OFFICE_NOTE).strip()
        base = dict(category=cat, name=v["name"], description=v["description"],
                    phone=v["phone"] if ok else "", email=v["email"] if ok else "",
                    url=v["url"], opening_times=v["opening_times"] if ok else "",
                    notes=notes, source_url=v["source_url"],
                    last_checked=v["date_checked"], publish="yes", verified="yes",
                    contact_hidden="", notes_detail=v["notes_detail"])
        for c in dict.fromkeys(councils):
            to_add.append((c, base))

    # ------------------------------------------------ report
    print(f"resources.json: {len(recs)} records, generated {data.get('generated')}")
    if errors:
        print(f"\nPROBLEMS ON THE SHEET ({len(errors)}) - these rows were skipped:")
        for e in errors:
            print("  -", e)
    print(f"\nEDITS TO EXISTING RECORDS ({len(edits)}):")
    for rid, name, ch in edits:
        print(f"  {rid} {name}")
        for k, (o, n) in ch.items():
            print(f"      {k}: {o!r} -> {n!r}")
    if missing:
        print(f"\nIN resources.json BUT NOT ON THE SHEET ({len(missing)}) - left untouched, not deleted:")
        print("  " + ", ".join(missing))
    if info:
        print("\nCOUNCIL NAMES READ AS:")
        for i in info:
            print("  -", i)
    print(f"\nNOT ADDED ({len(not_added)}):")
    for n in not_added:
        print("  -", n)
    print(f"\nNEW RECORDS ({len(to_add)}):")
    for c, b in to_add:
        print(f"  {b['category']:<10} {c:<22} {b['name']}")

    if not a.apply:
        print("\nReport only. Nothing written. Re-run with --apply to write these changes.")
        return

    # ------------------------------------------------ apply
    for rid, _, ch in edits:
        r = recs[rid]
        for k, (old, new) in ch.items():
            if k == "notes_list":
                if new:
                    r[k] = new
                else:
                    r.pop(k, None)
            elif new == "" and k not in r:
                continue
            elif new == "" and k in ("local_authority",):
                r.pop("local_authority", None)
                r.pop("la_code", None)
            else:
                r[k] = new
        if r.get("scope") == "local":
            r["la_code"] = la_code[r["local_authority"]]

    last = max(int(i[1:]) for i in recs)
    if os.path.exists(LAST_ID_FILE):
        last = max(last, int(open(LAST_ID_FILE).read().strip().lstrip("r")))
    for c, b in to_add:
        last += 1
        r = dict(b, id=f"r{last:04d}")
        if c == NATIONAL:
            r["scope"] = "national"
        else:
            r.update(scope="local", la_code=la_code[c], local_authority=c)
        data["resources"].append({k: r[k] for k in ORDER if k in r})
    with open(LAST_ID_FILE, "w") as f:
        f.write(f"r{last:04d}\n")

    if edits or to_add:
        prev = data.get("generated", "")
        stamp = datetime.now().replace(microsecond=0)
        try:
            prev_dt = datetime.strptime(prev, "%Y-%m-%dT%H:%M:%S")
            stamp = max(stamp, prev_dt + timedelta(minutes=1))
        except ValueError:
            pass
        data["generated"] = stamp.strftime("%Y-%m-%dT%H:%M:%S")
    with open(a.resources, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\nWROTE {a.resources}: {len(edits)} edited, {len(to_add)} added, "
          f"{len(data['resources'])} records, generated {data['generated']}. "
          f"Re-export the sheet so it picks up the new IDs.")


if __name__ == "__main__":
    main()
