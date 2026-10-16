# Working sheet for the KRISIS directory

`krisis-directory.xlsx` is built from `resources.json` and is meant to be opened in Google Sheets.

## Tabs
- **Directory**: one row for every record in `resources.json`. Correct mistakes here. Never edit the ID column.
- **Add new**: **one row per service.** Put every council it serves in *Councils served*, separated by `;` (for example `Fife; West Lothian`), or write `Scotland-wide`. Never copy a service into a separate row for each council.
- **Lists**: the values the dropdowns use. Leave this tab alone.

Hover over a column header to see the rules for that column. The source rules are the same as the verification protocol:
- Only the organisation's own website counts.
- Blank is better than a guess.
- Every row needs a source page and a date checked.

## Round trip
1. **Export**
   ```
   python3 tools/sheet_export.py resources.json sheet/krisis-directory.xlsx
   ```
   Upload the file to Google Drive, then open it with Google Sheets.
2. **Edit** the sheet in Google Sheets.
3. **Download** it with File > Download > Microsoft Excel (.xlsx).
4. **Check** what would change. This writes nothing:
   ```
   python3 tools/sheet_import.py resources.json ~/Downloads/krisis-directory.xlsx
   ```
5. **Apply** the changes, then re-export so the sheet shows the new IDs:
   ```
   python3 tools/sheet_import.py resources.json ~/Downloads/krisis-directory.xlsx --apply
   ```

## What the importer will not do
- Delete records. A row missing from the sheet is reported, and the record is left in place. Unpublish with Publish = `no` instead.
- Add an *Add new* row whose Check result is `dead`, `moved`, `unclear` or blank.
- Guess anything. An unknown council or category, a missing source page or date, a malformed phone or email, or a phone with no Phone type are all reported under NOT ADDED.
- Create a duplicate. A name already in `resources.json` is refused; correct that record on the Directory tab instead.
- Reuse an ID. New IDs continue from `tools/last_id.txt`, the highest ID ever issued, so deleted IDs are never handed out again.

The importer refuses a sheet that was exported before the last change to `resources.json`, because importing it would undo that change. If that happens, export a fresh sheet and copy your *Add new* rows into it.
