# CNC Setter 2.0

A web application that reads  CMM inspection reports and HyperMill NC programs, then recommends which tool offsets to adjust and by how much.

---

## What it does

1. You upload a CMM PDF (image-based  report) and one or more NC program files.
2. The parser OCR-reads every measured feature from the CMM report.
3. It parses the NC files to extract every machining operation and the tool that ran it.
4. It correlates CMM features to NC tools by matching diameters and operation types.
5. For each out-of-tolerance feature, it calculates the required tool offset correction and rates its confidence (HIGH / MEDIUM / LOW).
6. Results are saved and viewable in a split-pane browser UI, optionally alongside the part drawing PDF.

---

## Project structure

```
cnc-setter2.0/
├── generate_report (4).py   # Core parser & correlator (all algorithms here)
├── app.py                   # Flask web application (~450 lines)
├── templates/
│   ├── base.html            # Dark-theme Bootstrap 5 layout
│   ├── login.html           # Login page
│   ├── register.html        # Registration page
│   ├── dashboard.html       # Report list (card grid)
│   ├── upload.html          # File upload form
│   └── report.html          # Split-pane report viewer
├── uploads/                 # Uploaded files stored here (auto-created)
├── cncsetter.db             # SQLite database (auto-created, gitignored)
└── .gitignore
```

---

## Running the app

### Prerequisites

```bash
pip install flask pymupdf pytesseract pillow
apt install tesseract-ocr        # Ubuntu/Debian
# brew install tesseract         # macOS
```

### Start the server

```bash
python3 app.py
# → http://127.0.0.1:5000
```

First run creates the SQLite database automatically. Register the first user account on the `/register` page.

---

## Algorithm overview

### 1. CMM PDF parsing (`parse_cmm_pdf`)

- Opens the PDF with PyMuPDF
- Tries direct text extraction first; falls back to Tesseract OCR page-by-page
- Detects feature blocks by name lines matching `CIR`, `PLN`, `CYL`, `PNT`, `LIN`, `SET` prefixes
- Parses the data row: `nominal | +tol | -tol | meas | dev | outtol`
- Runs `_sanitize_meas()` on every measurement:
  - **Sign-flip fix**: if `nominal < 0` and `meas > 0` with similar magnitude, negates the measurement (common Tesseract OCR error)
  - **Implausibility check**: if `|dev| > max(tol_band × 100, 50 mm)`, marks the feature `suspect=True` and excludes it from matching

### 2. NC file parsing (`parse_nc_file`)

Two passes over each NC file:

**Pass 1 — tool registry**: pre-scan all `( TOOL NNN: description )` comment lines to build a `tool_num → description` map. This is necessary because HyperMill writes all tool headers at the top of the file, before any operation code.

**Pass 2 — operation extraction**:
- Tracks current operation from `( OPERATION NNN: ... )` comments
- Extracts tool number from `TNNN M6` lines and looks up description from the registry
- Captures canned cycles: G83 (peck drill), G85 (boring/reaming) with X/Y/Z/Q values
- Captures OPERATION descriptions that contain an explicit machined diameter (e.g. `CIRCULAR MILLING Ø158.400`) → creates a virtual `MILL_CIRCLE` entry
- Captures G3/G2 arc moves with I/J centre offsets when G41/G42 cutter compensation is active → radius = √(I²+J²) → diameter = 2×radius → another `MILL_CIRCLE` entry (only arcs ≥ Ø20 mm and appearing ≥ 2 times per operation)

### 3. Correlation (`correlate`)

For each out-of-tolerance (and non-suspect) CMM feature, `_find_dia_match()` tries four strategies in order:

| Priority | Strategy | Confidence |
|---|---|---|
| 1 | Exact reamer/bore tool diameter within 0.1 mm | HIGH |
| 2 | Drill diameter within 0.5 mm | HIGH |
| 3 | Any tool diameter within 0.5 mm (MEDIUM within 3 mm) | HIGH / MEDIUM |
| 4 | Circular milling operation diameter within 0.5 mm (3 mm) | HIGH / MEDIUM |
| — | No match | LOW |

- **Diameter (D-axis)** features are matched directly by nominal diameter. Cache key is `(feature_name, nominal)` so `CIR78` at Ø158.4 and Ø182 are matched independently.
- **Depth/position (Z, M, T-axis)** features use `_best_dia_for_name()`: find all diameter cache hits for the same feature name and reuse the best-confidence match.

### 4. Offset calculation

```
correction = nominal + tol_midpoint - meas
```

`tol_midpoint = (plus_tol - minus_tol) / 2`

For symmetric tolerances this equals `nominal - meas`. For asymmetric tolerances it targets the centre of the tolerance band rather than the nominal, which avoids biasing toward one side.

---

## Web app features

### Authentication
- Username/password login, bcrypt-equivalent SHA-256 hashing
- Session-based, all report routes require login
- First user registration is unrestricted; subsequent users can self-register (lock this down in production if needed)

### Dashboard
- Card grid of all saved reports
- Each card shows: part name, date, file name, total features, OOT count, HIGH/MEDIUM/LOW breakdown
- One-click delete (with confirmation)

### Upload
- Part name (optional label)
- CMM PDF (required)
- NC files — multiple allowed (`.nc`, `.cnc`, `.mpf`, `.h`, `.prg`)
- Drawing PDF (optional) — enables the side-by-side viewer

### Report viewer
- **Split-pane layout**: report on the left, drawing PDF in an iframe on the right (toggleable)
- **6 stat boxes**: total features, OOT count, suspect OCR, HIGH/MEDIUM/LOW match counts
- **Filter bar**: live text search, status filter (OOT/OK/Suspect), confidence filter, axis filter
- **Feature table**:
  - Tolerance bar (green/yellow/red fill showing deviation within tolerance band)
  - Colour-coded deviation: red = OOT positive, blue = OOT negative, green = in-tolerance
  - Tool override dropdown — pick any tool from the NC file; saved to DB and applied on future views
  - Offset correction value (monospace, colour-coded)
  - Match reason text
- **Inline part name editing**: click the name in the topbar to edit it
- **Drawing panel**: open/close with a button; PDF renders natively in browser

### API endpoints

| Method | Route | Description |
|---|---|---|
| POST | `/api/report/<id>/override` | Save a manual tool override for a feature |
| POST | `/api/report/<id>/name` | Rename a report's part name |

---

## Database schema

```sql
CREATE TABLE users (
    id            INTEGER PRIMARY KEY,
    username      TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    created_at    TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE reports (
    id               INTEGER PRIMARY KEY,
    user_id          INTEGER,
    part_name        TEXT,
    cmm_filename     TEXT,
    nc_filenames     TEXT,      -- JSON array of filenames
    drawing_filename TEXT,
    report_json      TEXT,      -- full report as JSON blob
    created_at       TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE tool_overrides (
    id          INTEGER PRIMARY KEY,
    report_id   INTEGER,
    feature_key TEXT,           -- "name:axis:nominal"
    tool        TEXT,
    tool_desc   TEXT,
    tool_type   TEXT
);
```

---

## Known limitations / future work

- **OCR accuracy**: Tesseract on image PDFs is imperfect. The sign-flip fix and suspect-flagging catch most errors, but edge cases exist. Better accuracy requires a higher-DPI scan or a CMM system that exports structured data.
- **Single-user uploads**: file upload processing is synchronous in the Flask request. For large CMM PDFs this can take 30–90 seconds. A task queue (Celery/RQ) would improve this.
- **No HTTPS / production hardening**: the default `SECRET_KEY` must be changed in production via the `SECRET_KEY` environment variable.
- **HyperMill format assumed**: the NC parser is tuned to HyperMill 2025 comment conventions. Other CAM systems use different comment styles and would need parser adjustments.
- **No multi-user isolation beyond login**: all users share the same upload directory; reports are filtered by `user_id` in queries but files are not encrypted.
