"""
CNCSETTER - HTML Report Generator
Usage:
    python generate_report.py cmm_report.pdf OP10.nc OP20.nc OP40.nc
    python generate_report.py cmm_report.pdf OP*.nc --output my_report.html

Outputs a single self-contained interactive HTML file with:
  - Tool offset recommendations (HIGH/MEDIUM/LOW confidence)
  - Tolerance bar visualization per feature
  - Filterable table (OOT only, by confidence, by axis type)
  - Click-to-expand detail panel per feature

Dependencies:
    pip install pymupdf pytesseract pillow
    apt install tesseract-ocr   (or brew install tesseract on macOS)
"""

import argparse
import json
import math
import os
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Tuple

# ─── DATA STRUCTURES ─────────────────────────────────────────────────────────

@dataclass
class CMMFeature:
    name: str
    feature_type: str   # CIR, PLN, CYL, PNT, LIN, SET
    axis: str           # D, R, F, M, Y, Z, T, A
    nominal: float
    plus_tol: float
    minus_tol: float
    meas: float
    dev: float
    outtol: float
    page: int

    @property
    def out_of_tol(self):
        return self.outtol > 0.0001

    @property
    def status(self):
        return "OOT" if self.out_of_tol else "OK"

    @property
    def tol_midpoint(self):
        """Signed offset from nominal to the centre of the tolerance band.
        Zero for symmetric tolerances; non-zero when plus_tol ≠ minus_tol."""
        return round((self.plus_tol - self.minus_tol) / 2, 4)

    @property
    def tol_used_pct(self):
        tol = self.plus_tol if self.dev >= 0 else self.minus_tol
        return round(abs(self.dev) / tol * 100, 1) if tol else 0

    def to_dict(self):
        return {
            "name": self.name, "type": self.feature_type, "axis": self.axis,
            "nominal": self.nominal, "plus_tol": self.plus_tol,
            "minus_tol": self.minus_tol,
            "tol_midpoint": self.tol_midpoint,
            "tol_center": round(self.nominal + self.tol_midpoint, 4),
            "meas": self.meas,
            "dev": round(self.dev, 4), "outtol": round(self.outtol, 4),
            "out_of_tol": self.out_of_tol, "status": self.status,
            "tol_used_pct": self.tol_used_pct, "page": self.page,
        }


@dataclass
class NCHole:
    op_file: str
    op_num: str
    desc: str
    tool: str
    tool_desc: str
    tool_type: str
    tool_dia: float
    cycle: str
    x: float
    y: float
    z: float


@dataclass
class OffsetRecommendation:
    feature: CMMFeature
    tool: str
    tool_desc: str
    op_num: str
    op_desc: str
    match_reason: str
    direction: str
    correction: float
    confidence: str
    n_holes: int = 0

    def to_dict(self):
        f = self.feature
        return {
            "feature": f.to_dict(),
            "tool": self.tool, "tool_desc": self.tool_desc,
            "op_num": self.op_num, "op_desc": self.op_desc,
            "match_reason": self.match_reason, "direction": self.direction,
            "correction": round(self.correction, 4),
            "correction_target": round(f.nominal + f.tol_midpoint, 4),
            "using_midpoint": f.tol_midpoint != 0.0,
            "confidence": self.confidence, "n_holes": self.n_holes,
        }


# ─── TOOL HELPERS ────────────────────────────────────────────────────────────

TOOL_KEYWORDS = {
    "REAMER": "REAMER", "REAMING": "REAMER", "H7": "REAMER", "H6": "REAMER",
    "DRILL": "DRILL", "DRILLING": "DRILL", "PECKING": "DRILL", "HELICAL": "DRILL",
    "BORE": "BORE", "BORING": "BORE",
    "THREAD": "THREAD", "THREADING": "THREAD",
    "CHAMFER": "CHAMFER", "CENTERING": "CHAMFER",
    "END MILL": "MILL", "ROUGHING": "MILL", "CONTOUR": "MILL",
    "BALL MILL": "MILL", "PLANE": "MILL", "FINISHING": "MILL",
}

def get_tool_type(text: str) -> str:
    t = (text or "").upper()
    for kw, val in TOOL_KEYWORDS.items():
        if kw in t:
            return val
    return "MILL"

def get_tool_dia(text: str) -> Optional[float]:
    m = re.search(r"\bD\s*(\d+\.?\d*)\b", text or "")
    return float(m.group(1)) if m else None


# ─── NC PARSER ───────────────────────────────────────────────────────────────

CANNED_CYCLE = re.compile(
    r"G98\s+(G8[1-9])\s+X([\d\.-]+)\s+Y([\d\.-]+)\s+Z([\d\.-]+)\s+R([\d\.-]+)"
)
CONTINUATION = re.compile(r"^X([\d\.-]+)\s+Y([\d\.-]+)$")


_RE_OP_DIA   = re.compile(r'\bD(\d+\.?\d*)\b')
_RE_OP_TOOL  = re.compile(r'\bT(\w+)\b')
_SKIP_OP_KWS = ("DRILL", "REAM", "CHAMFER", "THREAD", "PECKING", "PILOT",
                 "SPOT", "CENTER")


def parse_nc_file(filepath: str) -> Tuple[List[NCHole], List[dict]]:
    """Parse a single NC file. Returns (holes, ops_summary).

    In addition to canned-cycle holes (G83/G85), also creates virtual
    MILL_CIRCLE entries for milling operations whose OPERATION comment
    contains an explicit machined diameter (e.g. "T197 D158.4 FINISH").
    These are used to match large-bore CMM features that are machined
    by circular interpolation rather than canned cycles.
    """
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.read().splitlines()

    op_name = Path(filepath).stem
    holes: List[NCHole] = []
    ops_summary: List[dict] = []

    # Build tool registry from all TOOL comments (they appear at file top).
    # This avoids the stale cur_tool_desc bug where only the last TOOL comment
    # was remembered by the time each OPERATION comment was encountered.
    tool_registry: dict = {}   # tool_num → tool_desc
    for raw in lines:
        tm = re.match(r"\(\s*TOOL\s+(\S+):\s+(.+?)\s*\)", raw.strip())
        if tm:
            tool_registry[tm.group(1)] = tm.group(2).strip()

    cur_tool = None
    cur_tool_desc = None
    cur_op = None
    last_hole = None

    for line in lines:
        line = line.strip()

        m = re.match(r"T(\w+)\s+M6", line)
        if m:
            cur_tool = m.group(1)
            # Update cur_tool_desc from registry on tool change
            cur_tool_desc = tool_registry.get(cur_tool, cur_tool_desc)

        m = re.match(r"\(\s*TOOL\s+\S+:\s+(.+?)\s*\)", line)
        if m:
            cur_tool_desc = m.group(1).strip()

        m = re.match(r"\(\s*OPERATION\s+(\d+):\s+(.+?)\s*\)", line)
        if m:
            op_desc = m.group(2)
            # Resolve tool desc: try tool embedded in operation description
            mt = _RE_OP_TOOL.search(op_desc)
            op_tool = mt.group(1) if mt else cur_tool
            op_tool_desc = tool_registry.get(op_tool or "", "") or cur_tool_desc or ""
            cur_op = {
                "num": m.group(1), "desc": op_desc,
                "tool": op_tool, "tool_desc": op_tool_desc, "n_holes": 0,
            }
            ops_summary.append(cur_op)

        m = CANNED_CYCLE.match(line)
        if m and cur_op:
            dia = get_tool_dia(cur_tool_desc)
            desc_full = (cur_op["desc"] or "") + " " + (cur_tool_desc or "")
            h = NCHole(
                op_file=op_name, op_num=cur_op["num"], desc=cur_op["desc"],
                tool=cur_tool or "", tool_desc=cur_tool_desc or "",
                tool_type=get_tool_type(desc_full),
                tool_dia=dia or 0, cycle=m.group(1),
                x=float(m.group(2)), y=float(m.group(3)), z=float(m.group(4)),
            )
            holes.append(h)
            cur_op["n_holes"] += 1
            last_hole = h

        m = CONTINUATION.match(line)
        if m and last_hole and cur_op:
            h = NCHole(
                op_file=last_hole.op_file, op_num=last_hole.op_num,
                desc=last_hole.desc, tool=last_hole.tool,
                tool_desc=last_hole.tool_desc, tool_type=last_hole.tool_type,
                tool_dia=last_hole.tool_dia, cycle=last_hole.cycle,
                x=float(m.group(1)), y=float(m.group(2)), z=last_hole.z,
            )
            holes.append(h)
            cur_op["n_holes"] += 1
            last_hole = h

    # ── Circular milling ops: create virtual MILL_CIRCLE entries ─────────
    # HyperMill sometimes encodes the machined bore diameter directly in the
    # OPERATION comment, e.g. "T197 D158.4 FINISH".  Capture these so the
    # correlation engine can match them to large-bore CMM circle features.
    for op in ops_summary:
        desc = op.get("desc", "")
        m_dia = _RE_OP_DIA.search(desc)
        if not m_dia:
            continue
        machined_dia = float(m_dia.group(1))
        if machined_dia <= 0:
            continue
        # Skip operations already covered by canned cycles (drill/ream/etc.)
        desc_upper = desc.upper()
        if any(kw in desc_upper for kw in _SKIP_OP_KWS):
            continue
        tool_str = op.get("tool") or ""
        tool_desc_str = op.get("tool_desc") or ""
        holes.append(NCHole(
            op_file=op_name, op_num=op["num"], desc=desc,
            tool=tool_str, tool_desc=tool_desc_str,
            tool_type="MILL",
            tool_dia=machined_dia, cycle="MILL_CIRCLE",
            x=0.0, y=0.0, z=0.0,
        ))

    # ── Circular milling from G3/G2 arc radii ────────────────────────────
    # When G41/G42 cutter compensation is active (standard in HyperMill),
    # the programmed G3/G2 arc radius equals the bore radius of the machined
    # circle.  Machined diameter = 2 × arc radius.
    #
    # Strategy: for each operation, count how many times each arc radius
    # appears.  A radius that appears ≥2 times is a bore pass (helical
    # milling repeats the same radius many times); a one-off arc is just a
    # lead-in or corner.  Skip radii < 10mm (they're contour corners, not
    # bores we would measure with CMM).
    RE_G23 = re.compile(
        r"^G[23]\s+.*?I\s*([-\d.]+)\s+J\s*([-\d.]+)", re.IGNORECASE
    )
    # Collect arc diameters per operation
    op_arc_dias: dict = {op["num"]: Counter() for op in ops_summary}
    cur_op_num_arc = None
    for raw in lines:
        raw = raw.strip()
        mm = re.match(r"\(\s*OPERATION\s+(\d+):", raw)
        if mm:
            cur_op_num_arc = mm.group(1)
        if cur_op_num_arc:
            mm2 = RE_G23.match(raw)
            if mm2:
                I = float(mm2.group(1))
                J = float(mm2.group(2))
                r = math.sqrt(I * I + J * J)
                if r >= 10.0:  # exclude tiny arcs (lead-ins, corner radii)
                    dia = round(r * 2, 3)
                    op_arc_dias[cur_op_num_arc][dia] += 1

    # Diameters already captured via explicit OPERATION description
    already_captured_dias = {h.tool_dia for h in holes if h.cycle == "MILL_CIRCLE"}

    for op in ops_summary:
        arc_counter = op_arc_dias.get(op["num"], Counter())
        tool_str = op.get("tool") or ""
        tool_desc_str = op.get("tool_desc") or ""
        desc = op.get("desc", "")
        for dia, cnt in arc_counter.items():
            if cnt < 2:
                continue  # single arc = lead-in, not a bore pass
            # Skip if already covered by an explicit-diameter MILL_CIRCLE entry
            # for this same operation (avoid double-entry at the same diameter)
            if any(
                h.op_num == op["num"] and abs(h.tool_dia - dia) < 0.01
                for h in holes
                if h.cycle == "MILL_CIRCLE"
            ):
                continue
            holes.append(NCHole(
                op_file=op_name, op_num=op["num"], desc=desc,
                tool=tool_str, tool_desc=tool_desc_str,
                tool_type="MILL",
                tool_dia=dia, cycle="MILL_CIRCLE",
                x=0.0, y=0.0, z=0.0,
            ))

    return holes, ops_summary


# ─── CMM PARSER (OCR) ────────────────────────────────────────────────────────

RE_DATA_ROW = re.compile(
    r"^([A-Z]{1,4}\s*[A-Z0-9]*)\s+([-\d]+\.?\d*)\s+([\d]+\.?\d*)\s+([-\d]*\.?\d*)"
    r"\s+([-\d]+\.?\d*)\s+([-\d]+\.?\d*)\s+([\d]+\.?\d*)"
)
RE_DIAM_ROW = re.compile(
    r"^D\s+([\d]+\.?\d*)\s+([\d]+\.?\d*)\s+([\d]+\.?\d*)"
    r"\s+([\d]+\.?\d*)\s+([-\d]+\.?\d*)\s+([\d]+\.?\d*)"
)
RE_AXIS_ROW = re.compile(
    r"^([MYRZTF])\s+([-\d]+\.?\d*)\s+([\d]+\.?\d*)\s+([\d]+\.?\d*)"
    r"\s+([-\d]+\.?\d*)\s+([-\d]+\.?\d*)\s+([\d]+\.?\d*)"
)
RE_FEAT_HEADER = re.compile(
    r"(?:^|\s)(\d+\s*[-–]\s*(CIR|PLN|CYL|LIN|PNT|SET)\s*[A-Z0-9]*)", re.IGNORECASE
)
RE_FEAT_NAME = re.compile(r"\b(CIR|PLN|CYL|LIN|PNT|SET)\s*([A-Z0-9]+)\b", re.IGNORECASE)


def _compute_outtol(dev: float, plus_tol: float, minus_tol: float) -> float:
    """Return the out-of-tolerance amount (≥0) computed purely from the
    deviation and tolerance values, independent of the OCR'd outtol column."""
    if dev > plus_tol + 1e-9:
        return round(dev - plus_tol, 6)
    if dev < -(minus_tol + 1e-9):
        return round(-dev - minus_tol, 6)
    return 0.0


def parse_cmm_pdf(filepath: str, verbose: bool = False) -> List[CMMFeature]:
    """
    Parse a Calypso CMM PDF.

    Extraction strategy (tried in order for each page):
      1. Direct text extraction via PyMuPDF (fast, reliable for vector PDFs
         — the typical Calypso "Print to PDF" output).
      2. OCR via Tesseract on any embedded raster images (fallback for
         scanned / image-based PDFs).

    Requires: pymupdf (fitz); for OCR fallback also pytesseract + Pillow +
    tesseract-ocr.
    """
    try:
        import fitz
    except ImportError:
        print("[ERROR] pymupdf not installed.  Run: pip install pymupdf")
        sys.exit(1)

    # pytesseract and Pillow are only needed for the OCR fallback path.
    # We import them lazily below so the tool works without them when the
    # PDF has selectable text (the common Calypso case).
    try:
        import pytesseract
        from PIL import Image
        import io
        _ocr_available = True
    except ImportError:
        pytesseract = None          # type: ignore[assignment]
        Image = None                # type: ignore[assignment]
        io = None                   # type: ignore[assignment]
        _ocr_available = False

    # Windows: if OCR is available but tesseract isn't on PATH, find it.
    if _ocr_available and sys.platform == "win32":
        import shutil
        if not shutil.which("tesseract"):
            for _path in [
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            ]:
                if os.path.isfile(_path):
                    pytesseract.pytesseract.tesseract_cmd = _path
                    break

    try:
        from tqdm import tqdm
    except ImportError:
        # tqdm is optional — fall back to plain range
        tqdm = None

    doc = fitz.open(filepath)
    features: List[CMMFeature] = []
    cur_name = None
    cur_type = None
    total_pages = len(doc)

    page_iter = range(total_pages)
    if verbose:
        if tqdm:
            page_iter = tqdm(
                page_iter,
                total=total_pages,
                desc="  Parse",
                unit="pg",
                bar_format=(
                    "  Parse [{bar:30}] {n_fmt}/{total_fmt} pages  "
                    "({elapsed} elapsed, ~{remaining} left)"
                ),
                ncols=72,
            )
        else:
            print(f"  Parse: 0/{total_pages} pages", end="\r", flush=True)

    for page_num in page_iter:
        if verbose and not tqdm:
            print(f"  Parse: {page_num + 1}/{total_pages} pages", end="\r", flush=True)

        page = doc[page_num]

        # ── Collect text-line sources for this page ───────────────────────
        # Strategy: try direct vector-text extraction first (Calypso normally
        # produces a selectable-text PDF, not an image PDF).  Fall back to
        # per-image OCR only when the page has no extractable text (scanned /
        # pure-image PDF).
        page_sources: List[List[str]] = []

        # -- Direct text extraction (vector PDF) -------------------------
        try:
            raw = page.get_text("rawdict", flags=0)
            words = []
            for blk in raw.get("blocks", []):
                if blk.get("type") != 0:          # skip image blocks
                    continue
                for ln in blk.get("lines", []):
                    for span in ln.get("spans", []):
                        txt = span.get("text", "").strip()
                        if not txt:
                            continue
                        bbox = span.get("bbox", [0, 0, 0, 0])
                        y_mid = (bbox[1] + bbox[3]) / 2
                        page_sources  # ensure closure captures outer list
                        words.append((y_mid, bbox[0], txt))
            if len(words) > 8:
                # Group words by Y position (within 4 pt) to reconstruct rows.
                words.sort(key=lambda w: (w[0], w[1]))
                rows: List[str] = []
                cur_y: Optional[float] = None
                cur_row: List[Tuple[float, str]] = []
                for y, x, txt in words:
                    if cur_y is None or abs(y - cur_y) <= 4.0:
                        cur_row.append((x, txt))
                        if cur_y is None:
                            cur_y = y
                    else:
                        cur_row.sort(key=lambda w: w[0])
                        rows.append(" ".join(t for _, t in cur_row))
                        cur_row = [(x, txt)]
                        cur_y = y
                if cur_row:
                    cur_row.sort(key=lambda w: w[0])
                    rows.append(" ".join(t for _, t in cur_row))
                direct_lines = [r.strip() for r in rows if r.strip()]
                if direct_lines:
                    page_sources.append(direct_lines)
        except Exception:
            pass

        # -- OCR fallback (image-based / scanned PDF) --------------------
        if not page_sources and _ocr_available:
            for img_info in page.get_images(full=True):
                try:
                    xref = img_info[0]
                    base_image = doc.extract_image(xref)
                    img = Image.open(io.BytesIO(base_image["image"]))
                    w, h = img.size
                    if w > 1200:
                        scale = 1200 / w
                        img = img.resize((int(w * scale), int(h * scale)),
                                         Image.LANCZOS)
                    text = pytesseract.image_to_string(img,
                                                       config="--psm 6 --oem 3")
                    ocr_lines = [ln.strip() for ln in text.split("\n")]
                    page_sources.append(ocr_lines)
                except Exception:
                    pass

        # ── Two-pass processing — identical for both direct and OCR lines ─
        for img_lines in page_sources:

            # ── Pass 1: locate every feature-name declaration in this block ──
            # Stores (line_index, name, type) so that D/axis rows can look up
            # the *nearest preceding* name in the same block instead of relying
            # on cur_name which may come from the previous image.
            img_name_locs: List[Tuple[int, str, str]] = []
            for i, line in enumerate(img_lines):
                if not line:
                    continue
                # Explicit header: "9 - CIR84", "KEY5 - CYL F"
                m = RE_FEAT_HEADER.search(line)
                if m:
                    raw2 = m.group(1)
                    type_m = re.search(r"(CIR|PLN|CYL|LIN|PNT|SET)", raw2, re.I)
                    if type_m:
                        ftype = type_m.group(1).upper()
                        after = raw2.split("-")[-1].strip()
                        name_m = RE_FEAT_NAME.search(after)
                        fname = (
                            name_m.group(1).upper() + name_m.group(2).upper()
                            if name_m
                            else re.sub(r"[^A-Z0-9]", "", after.upper()) or ftype
                        )
                        img_name_locs.append((i, fname, ftype))
                    continue
                # Inline name on a non-data line: "CIR84" or "CIR 84" alone
                if (not RE_DIAM_ROW.match(line)
                        and not RE_AXIS_ROW.match(line)
                        and not RE_DATA_ROW.match(line)):
                    m2 = RE_FEAT_NAME.search(line)
                    if m2:
                        img_name_locs.append((
                            i,
                            m2.group(1).upper() + m2.group(2).upper(),
                            m2.group(1).upper(),
                        ))

            # After processing this block, advance the global cursors so the
            # *next* block can fall back to the last name seen here.
            if img_name_locs:
                _, cur_name, cur_type = img_name_locs[-1]

            def nearest_img_name(line_idx: int) -> Tuple[Optional[str], Optional[str]]:
                """(name, type) of the closest preceding name declaration in
                this block; falls forward if nothing precedes the row."""
                for li, name, typ in reversed(img_name_locs):
                    if li <= line_idx:
                        return name, typ
                # Nothing before this line — try the first one after
                for li, name, typ in img_name_locs:
                    if li > line_idx:
                        return name, typ
                return None, None

            # ── Pass 2: extract measurement rows ─────────────────────────────
            for i, line in enumerate(img_lines):
                if not line:
                    continue

                # Diameter row (check BEFORE RE_DATA_ROW — "D" is also matched
                # by the generic pattern): "D 152.400 0.075 0.060 152.384 ..."
                m = RE_DIAM_ROW.match(line)
                if m:
                    local_name, local_type = nearest_img_name(i)
                    name  = local_name or cur_name
                    ftype = local_type or cur_type or "CIR"
                    if name:
                        try:
                            nominal = float(m.group(1))
                            plus_t  = float(m.group(2))
                            minus_t = float(m.group(3))
                            meas    = float(m.group(4))
                            dev     = round(meas - nominal, 6)
                            features.append(CMMFeature(
                                name=name, feature_type=ftype, axis="D",
                                nominal=nominal, plus_tol=plus_t, minus_tol=minus_t,
                                meas=meas, dev=dev,
                                outtol=_compute_outtol(dev, plus_t, minus_t),
                                page=page_num + 1,
                            ))
                        except ValueError:
                            pass
                    continue

                # Axis row (check BEFORE RE_DATA_ROW — single letters like
                # M/Y/Z/T are also caught by the generic pattern):
                # "M 11.900 0.030 0.030 11.520 -0.380 0.350"
                m = RE_AXIS_ROW.match(line)
                if m:
                    local_name, local_type = nearest_img_name(i)
                    name  = local_name or cur_name
                    ftype = local_type or cur_type or "LIN"
                    if name:
                        try:
                            nominal = float(m.group(2))
                            plus_t  = float(m.group(3))
                            minus_t = float(m.group(4))
                            meas    = float(m.group(5))
                            dev     = round(meas - nominal, 6)
                            features.append(CMMFeature(
                                name=name, feature_type=ftype, axis=m.group(1),
                                nominal=nominal, plus_tol=plus_t, minus_tol=minus_t,
                                meas=meas, dev=dev,
                                outtol=_compute_outtol(dev, plus_t, minus_t),
                                page=page_num + 1,
                            ))
                        except ValueError:
                            pass
                    continue

                # Full feature row (name embedded) — fallback after specific
                # D/axis checks: "CIR85 0.000 0.030 0.032 0.032 0.002 0.000"
                m = RE_DATA_ROW.match(line)
                if m:
                    fname = m.group(1).replace(" ", "").upper()
                    type_m = re.match(r"(CIR|PLN|CYL|LIN|PNT|SET)", fname, re.I)
                    if type_m:
                        try:
                            nominal = float(m.group(2))
                            plus_t  = float(m.group(3))
                            minus_t = float(m.group(4)) if m.group(4) else plus_t
                            meas    = float(m.group(5))
                            dev     = round(meas - nominal, 6)   # computed, not OCR
                            features.append(CMMFeature(
                                name=fname, feature_type=type_m.group(1).upper(),
                                axis="F",
                                nominal=nominal, plus_tol=plus_t, minus_tol=minus_t,
                                meas=meas, dev=dev,
                                outtol=_compute_outtol(dev, plus_t, minus_t),
                                page=page_num + 1,
                            ))
                            # The embedded name also becomes the context for
                            # any following D/axis rows in this block.
                            cur_name = fname
                            cur_type = type_m.group(1).upper()
                        except ValueError:
                            pass
                    continue

    if verbose and not tqdm:
        print()  # newline after fallback progress line

    return features


# ─── CORRELATION ENGINE ──────────────────────────────────────────────────────

def _find_dia_match(
    feat: "CMMFeature",
    nc_holes: List[NCHole],
) -> Tuple[Optional[NCHole], str, str]:
    """Return (best_hole, confidence, reason) for a diameter CMM feature."""
    # 1. Reamer — tightest fit, highest confidence
    for h in nc_holes:
        if h.tool_type == "REAMER" and h.tool_dia > 0:
            if abs(h.tool_dia - feat.nominal) < 0.15:
                return h, "HIGH", (
                    f"Reamer T{h.tool} D{h.tool_dia:.3f}mm → nominal Ø{feat.nominal:.3f}mm"
                )

    # 2. Drill — loose fit, medium confidence
    for h in nc_holes:
        if h.tool_type == "DRILL" and h.tool_dia > 0:
            if abs(h.tool_dia - feat.nominal) < 0.2:
                return h, "MEDIUM", (
                    f"Drill T{h.tool} D{h.tool_dia:.3f}mm → nominal Ø{feat.nominal:.3f}mm"
                )

    # 3. Boring bar — wider match band for large bores
    for h in nc_holes:
        if h.tool_type == "BORE" and h.tool_dia > 0:
            if abs(h.tool_dia - feat.nominal) < 0.5:
                return h, "MEDIUM", (
                    f"Boring bar T{h.tool} D{h.tool_dia:.3f}mm → nominal Ø{feat.nominal:.3f}mm"
                )

    # 4. Circular milling op — diameter encoded in operation description
    #    (HyperMill writes e.g. "T197 D158.4 FINISH" for a bore pass)
    best_mill: Optional[NCHole] = None
    best_mill_diff = float("inf")
    for h in nc_holes:
        if h.cycle == "MILL_CIRCLE" and h.tool_dia > 0:
            diff = abs(h.tool_dia - feat.nominal)
            if diff < best_mill_diff:
                best_mill_diff = diff
                best_mill = h
    if best_mill is not None:
        if best_mill_diff < 0.5:
            return best_mill, "HIGH", (
                f"Circular mill T{best_mill.tool} machined "
                f"Ø{best_mill.tool_dia:.3f}mm → nominal Ø{feat.nominal:.3f}mm"
            )
        if best_mill_diff < 3.0:
            return best_mill, "MEDIUM", (
                f"Nearest circular mill T{best_mill.tool} "
                f"Ø{best_mill.tool_dia:.3f}mm (Δ{best_mill_diff:.2f}mm) "
                f"→ nominal Ø{feat.nominal:.3f}mm"
            )

    # 5. No match
    if feat.nominal > 15:
        msg = f"Large bore Ø{feat.nominal:.3f}mm — assign manually (milling/boring op)"
    else:
        msg = f"No matching tool for Ø{feat.nominal:.3f}mm"
    return None, "LOW", msg


def correlate(
    cmm_features: List[CMMFeature],
    nc_holes: List[NCHole],
) -> List[OffsetRecommendation]:
    """Match CMM features to NC tools and build offset recommendations.

    Correction is always computed relative to the *centre of the tolerance
    band*, not to the bare nominal.  For symmetric tolerances this is
    identical to the classic  correction = -dev  formula.  For asymmetric
    (bilateral) tolerances the correction steers the dimension to the
    midpoint of the band, giving maximum margin on both sides.

        tol_midpoint = (plus_tol - minus_tol) / 2   # offset from nominal
        correction   = tol_midpoint - dev
    """

    # ── Pass 1: resolve each diameter feature → best NC tool ─────────────
    # Key: (name, nominal) so that a feature name that appears at multiple
    # diameters (e.g. CIR78 at Ø158.4 AND Ø182.0) each gets its own match.
    dia_cache: dict = {}   # (name, nominal) → (NCHole|None, confidence, reason)
    for feat in cmm_features:
        key = (feat.name, feat.nominal)
        if feat.axis == "D" and key not in dia_cache:
            dia_cache[key] = _find_dia_match(feat, nc_holes)

    def _best_dia_for_name(name: str):
        """Return the best (hole, confidence, reason) cached for any nominal
        of *name*, preferring HIGH > MEDIUM > LOW, then smallest LOW nominal."""
        candidates = [v for (n, _), v in dia_cache.items() if n == name]
        if not candidates:
            return None, "LOW", ""
        order = {"HIGH": 0, "MEDIUM": 1, "LOW": 2}
        candidates.sort(key=lambda v: order.get(v[1], 9))
        return candidates[0]

    # ── Pass 2: build recommendations for every feature with a deviation ──
    recs: List[OffsetRecommendation] = []

    for feat in cmm_features:
        if abs(feat.dev) < 0.0001:
            continue

        # Correction targets the tolerance band centre (midpoint).
        # When plus_tol == minus_tol, tol_midpoint == 0 → same as -dev.
        tol_midpoint = feat.tol_midpoint
        correction = tol_midpoint - feat.dev

        matched: Optional[NCHole] = None
        reason = ""
        confidence = "LOW"

        if feat.axis == "D":
            matched, confidence, reason = dia_cache.get(
                (feat.name, feat.nominal),
                (None, "LOW", f"No matching tool for Ø{feat.nominal:.3f}mm"),
            )
            direction = "RADIAL (diameter)"

        elif feat.axis in ("Z", "M", "T"):
            direction = "AXIAL"
            # Inherit the NC operation from the best diameter match for this
            # feature name.  Works even when multiple bores share a name.
            cached_hole, cached_conf, _ = _best_dia_for_name(feat.name)
            if cached_hole is not None:
                matched = cached_hole
                confidence = "MEDIUM"
                reason = (
                    f"{feat.axis}-axis on {feat.feature_type} {feat.name} "
                    f"linked via diameter match → T{matched.tool}"
                )
            else:
                reason = f"{feat.axis}-axis deviation on {feat.feature_type} {feat.name}"

        else:
            direction = "POSITIONAL"
            reason = f"{feat.axis}-axis deviation on {feat.feature_type} {feat.name}"

        n_holes = 0
        if matched:
            n_holes = sum(
                1 for h in nc_holes
                if h.op_num == matched.op_num and h.op_file == matched.op_file
            )

        recs.append(OffsetRecommendation(
            feature=feat,
            tool=matched.tool if matched else "—",
            tool_desc=matched.tool_desc if matched else "—",
            op_num=matched.op_num if matched else "—",
            op_desc=matched.desc if matched else "—",
            match_reason=reason,
            direction=direction,
            correction=correction,
            confidence=confidence,
            n_holes=n_holes,
        ))

    return sorted(recs, key=lambda r: (-r.feature.out_of_tol, -abs(r.feature.dev)))


# ─── HTML GENERATOR ──────────────────────────────────────────────────────────

def generate_html(
    recs: List[OffsetRecommendation],
    cmm_features: List[CMMFeature],
    nc_holes: List[NCHole],
    ops_summary: List[dict],
    part_name: str = "",
    cmm_filename: str = "",
    nc_filenames: List[str] = None,
) -> str:
    """Generate a self-contained interactive HTML report."""

    nc_filenames = nc_filenames or []
    oot_features = [f for f in cmm_features if f.out_of_tol]
    ok_features  = [f for f in cmm_features if not f.out_of_tol]
    recs_high    = sum(1 for r in recs if r.confidence == "HIGH" and r.feature.out_of_tol)
    recs_medium  = sum(1 for r in recs if r.confidence == "MEDIUM" and r.feature.out_of_tol)
    max_dev      = max((abs(f.dev) for f in cmm_features), default=0)
    max_dev_feat = next((f for f in cmm_features if abs(f.dev) == max_dev), None)

    recs_json = json.dumps([r.to_dict() for r in recs], indent=None)
    nc_files_str = " · ".join(nc_filenames) or "—"

    # Build ops table rows
    ops_rows = ""
    for op in ops_summary:
        n = op.get("n_holes", 0)
        ops_rows += f"""<tr>
          <td>OP{op['num']}</td>
          <td>{op['desc']}</td>
          <td>T{op['tool'] or '—'}</td>
          <td style="color:var(--dim)">{(op['tool_desc'] or '')[:60]}</td>
          <td style="text-align:right;color:{'var(--accent)' if n>0 else 'var(--dim)'}">{n}</td>
        </tr>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CNCSETTER — {part_name}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@300;400;500;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
:root{{
  --bg:#0d0f12;--surface:#141720;--border:#232733;--border2:#2e3347;
  --text:#c8cdd8;--dim:#5a6075;--accent:#00d4ff;--ok:#00c87a;
  --oot:#ff4757;--warn:#ffa502;--high:#00c87a;--medium:#ffa502;--low:#ff6b35;
  --mono:'IBM Plex Mono',monospace;--sans:'IBM Plex Sans',sans-serif;
}}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:var(--sans);background:var(--bg);color:var(--text);min-height:100vh;font-size:13px}}
header{{display:flex;align-items:center;justify-content:space-between;padding:0 24px;height:52px;border-bottom:1px solid var(--border);background:var(--surface);position:sticky;top:0;z-index:100}}
.logo{{font-family:var(--mono);font-size:15px;font-weight:600;letter-spacing:.15em;color:var(--accent);display:flex;align-items:center;gap:10px}}
.logo-dot{{width:8px;height:8px;border-radius:50%;background:var(--accent);animation:pulse 2s ease-in-out infinite}}
@keyframes pulse{{0%,100%{{opacity:1}}50%{{opacity:.3}}}}
.part-badge{{font-family:var(--mono);font-size:11px;color:var(--dim);border:1px solid var(--border2);padding:3px 10px;border-radius:3px}}
.layout{{display:grid;grid-template-columns:280px 1fr;height:calc(100vh - 52px)}}
.sidebar{{border-right:1px solid var(--border);background:var(--surface);overflow-y:auto;display:flex;flex-direction:column}}
.sidebar-section{{padding:16px;border-bottom:1px solid var(--border)}}
.sidebar-label{{font-family:var(--mono);font-size:10px;font-weight:500;letter-spacing:.12em;color:var(--dim);text-transform:uppercase;margin-bottom:10px}}
.stats-grid{{display:grid;grid-template-columns:1fr 1fr;gap:8px}}
.stat-card{{background:var(--bg);border:1px solid var(--border);border-radius:4px;padding:10px}}
.stat-val{{font-family:var(--mono);font-size:22px;font-weight:600;line-height:1}}
.stat-val.oot{{color:var(--oot)}}.stat-val.ok{{color:var(--ok)}}.stat-val.acc{{color:var(--accent)}}
.stat-label{{font-size:10px;color:var(--dim);margin-top:4px;text-transform:uppercase;letter-spacing:.08em}}
.filter-row{{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:6px}}
.pill{{padding:4px 10px;border-radius:12px;border:1px solid var(--border2);font-family:var(--mono);font-size:10px;cursor:pointer;transition:all .15s;background:transparent;color:var(--dim)}}
.pill:hover{{border-color:var(--accent);color:var(--accent)}}
.pill.active{{background:var(--accent);border-color:var(--accent);color:#000;font-weight:600}}
.pill.oot-pill.active{{background:var(--oot);border-color:var(--oot);color:#fff}}
.pill.high-pill.active{{background:var(--high);border-color:var(--high);color:#000}}
.pill.low-pill.active{{background:var(--low);border-color:var(--low);color:#fff}}
.main{{overflow-y:auto;background:var(--bg)}}
.results{{padding:20px}}
.section-header{{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid var(--border)}}
.section-title{{font-family:var(--mono);font-size:11px;font-weight:500;letter-spacing:.12em;color:var(--dim);text-transform:uppercase}}
.section-count{{font-family:var(--mono);font-size:11px;color:var(--dim)}}
.summary-bar{{display:flex;gap:16px;padding:12px 20px;background:var(--surface);border-bottom:1px solid var(--border);flex-wrap:wrap}}
.sum-item{{display:flex;align-items:center;gap:6px;font-family:var(--mono);font-size:11px}}
.sum-dot{{width:6px;height:6px;border-radius:50%}}
.sum-dot.oot{{background:var(--oot)}}.sum-dot.ok{{background:var(--ok)}}.sum-dot.acc{{background:var(--accent)}}.sum-dot.high{{background:var(--high)}}
.col-header-row{{display:grid;grid-template-columns:3px 150px 95px 95px 95px 115px 1fr 95px;border-bottom:1px solid var(--border2);padding:6px 0;background:var(--surface);position:sticky;top:0;z-index:10}}
.col-header{{padding:0 10px;font-family:var(--mono);font-size:9px;font-weight:500;letter-spacing:.12em;color:var(--dim);text-transform:uppercase}}
.rec-row{{display:grid;grid-template-columns:3px 150px 95px 95px 95px 115px 1fr 95px;align-items:stretch;border-bottom:1px solid var(--border);transition:background .1s;cursor:pointer;min-height:48px}}
.rec-row:hover{{background:rgba(255,255,255,.025)}}
.rec-stripe{{width:3px;align-self:stretch}}
.rec-stripe.oot{{background:var(--oot)}}.rec-stripe.ok{{background:var(--ok)}}
.rec-cell{{padding:10px;display:flex;flex-direction:column;justify-content:center;gap:2px}}
.rec-feature-name{{font-family:var(--mono);font-size:12px;font-weight:600;color:var(--text)}}
.rec-feature-sub{{font-family:var(--mono);font-size:10px;color:var(--dim)}}
.val-main{{font-family:var(--mono);font-size:12px;color:var(--text)}}
.val-sub{{font-family:var(--mono);font-size:10px;color:var(--dim)}}
.deviation-val{{font-family:var(--mono);font-size:13px;font-weight:600}}
.deviation-val.pos{{color:var(--ok)}}.deviation-val.neg{{color:var(--oot)}}
.correction-val{{font-family:var(--mono);font-size:13px;font-weight:600;display:flex;align-items:center;gap:4px}}
.correction-val.up{{color:var(--ok)}}.correction-val.down{{color:var(--oot)}}
.bar-container{{display:flex;flex-direction:column;justify-content:center;gap:3px}}
.tol-bar-wrap{{width:100%;height:4px;background:var(--border2);border-radius:2px;overflow:hidden}}
.tol-bar{{height:100%;border-radius:2px;transition:width .3s}}
.tol-bar.ok{{background:var(--ok)}}.tol-bar.warn{{background:var(--warn)}}.tol-bar.oot{{background:var(--oot)}}
.conf-badge{{display:inline-flex;align-items:center;padding:2px 7px;border-radius:3px;font-family:var(--mono);font-size:10px;font-weight:500;letter-spacing:.05em}}
.conf-badge.HIGH{{background:rgba(0,200,122,.15);color:var(--high);border:1px solid rgba(0,200,122,.3)}}
.conf-badge.MEDIUM{{background:rgba(255,165,2,.15);color:var(--medium);border:1px solid rgba(255,165,2,.3)}}
.conf-badge.LOW{{background:rgba(255,107,53,.12);color:var(--low);border:1px solid rgba(255,107,53,.25)}}
.status-badge{{display:inline-flex;align-items:center;padding:2px 7px;border-radius:3px;font-family:var(--mono);font-size:10px;font-weight:600}}
.status-badge.OOT{{background:rgba(255,71,87,.15);color:var(--oot);border:1px solid rgba(255,71,87,.3)}}
.status-badge.OK{{background:rgba(0,200,122,.12);color:var(--ok);border:1px solid rgba(0,200,122,.25)}}
.tool-info{{display:flex;flex-direction:column;gap:3px}}
.tool-num{{font-family:var(--mono);font-size:11px;color:var(--accent);font-weight:500}}
.tool-desc-text{{font-size:11px;color:var(--dim);line-height:1.3;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:180px}}
.tool-assign-btn{{font-family:var(--mono);font-size:10px;color:var(--dim);border:1px dashed var(--border2);background:transparent;padding:3px 8px;border-radius:3px;cursor:pointer;transition:all .15s}}
.tool-assign-btn:hover{{border-color:var(--accent);color:var(--accent)}}
.detail-panel{{margin-top:16px;background:var(--surface);border:1px solid var(--border);border-radius:4px;overflow:hidden;animation:fadeIn .2s ease}}
@keyframes fadeIn{{from{{opacity:0;transform:translateY(-4px)}}to{{opacity:1;transform:translateY(0)}}}}
.detail-header{{display:flex;align-items:center;justify-content:space-between;padding:12px 16px;border-bottom:1px solid var(--border);background:rgba(0,212,255,.03)}}
.detail-title{{font-family:var(--mono);font-size:13px;font-weight:600;color:var(--accent)}}
.detail-close{{background:transparent;border:none;color:var(--dim);cursor:pointer;font-size:16px;line-height:1}}
.detail-close:hover{{color:var(--text)}}
.detail-body{{display:grid;grid-template-columns:1fr 1fr 1fr}}
.detail-group{{padding:14px 16px;border-right:1px solid var(--border)}}
.detail-group:last-child{{border-right:none}}
.detail-group-label{{font-family:var(--mono);font-size:9px;letter-spacing:.12em;color:var(--dim);text-transform:uppercase;margin-bottom:10px}}
.detail-row{{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px}}
.detail-key{{font-size:11px;color:var(--dim)}}
.detail-val{{font-family:var(--mono);font-size:11px;color:var(--text)}}
.offset-box{{margin-top:12px;padding:10px;background:rgba(0,212,255,.05);border:1px solid var(--border2);border-radius:3px;font-family:var(--mono);font-size:10px;color:var(--dim);line-height:1.8}}
.tabs{{display:flex;gap:0;border-bottom:1px solid var(--border);padding:0 20px;background:var(--surface)}}
.tab{{padding:10px 18px;font-family:var(--mono);font-size:11px;letter-spacing:.08em;color:var(--dim);cursor:pointer;border-bottom:2px solid transparent;transition:all .15s}}
.tab:hover{{color:var(--text)}}
.tab.active{{color:var(--accent);border-bottom-color:var(--accent)}}
.tab-content{{display:none}}.tab-content.active{{display:block}}
.ops-table{{width:100%;border-collapse:collapse;font-family:var(--mono);font-size:11px}}
.ops-table th{{text-align:left;padding:8px 12px;font-size:10px;letter-spacing:.1em;color:var(--dim);border-bottom:1px solid var(--border);text-transform:uppercase}}
.ops-table td{{padding:8px 12px;border-bottom:1px solid var(--border);color:var(--text)}}
.ops-table tr:hover td{{background:rgba(255,255,255,.02)}}
::-webkit-scrollbar{{width:6px;height:6px}}
::-webkit-scrollbar-track{{background:var(--bg)}}
::-webkit-scrollbar-thumb{{background:var(--border2);border-radius:3px}}
.toast{{position:fixed;bottom:20px;right:20px;background:var(--surface);border:1px solid var(--border2);border-left:3px solid var(--accent);padding:12px 16px;border-radius:4px;font-family:var(--mono);font-size:12px;color:var(--text);z-index:999;animation:slideIn .2s ease;max-width:300px}}
@keyframes slideIn{{from{{transform:translateX(20px);opacity:0}}}}
</style>
</head>
<body>

<header>
  <div class="logo"><div class="logo-dot"></div>CNCSETTER</div>
  <div style="display:flex;align-items:center;gap:10px">
    <div class="part-badge">{part_name} · {cmm_filename}</div>
  </div>
</header>

<div class="layout">
  <aside class="sidebar">
    <div class="sidebar-section">
      <div class="sidebar-label">CMM Summary</div>
      <div class="stats-grid">
        <div class="stat-card"><div class="stat-val oot">{len(oot_features)}</div><div class="stat-label">Out of Tol</div></div>
        <div class="stat-card"><div class="stat-val ok">{len(ok_features)}</div><div class="stat-label">In Tol</div></div>
        <div class="stat-card"><div class="stat-val acc">{len(nc_holes)}</div><div class="stat-label">NC Holes</div></div>
        <div class="stat-card"><div class="stat-val acc">{recs_high + recs_medium}</div><div class="stat-label">Auto-matched</div></div>
      </div>
    </div>
    <div class="sidebar-section">
      <div class="sidebar-label">Filter by status</div>
      <div class="filter-row">
        <button class="pill active" onclick="setFilter('all',this)">All</button>
        <button class="pill oot-pill" onclick="setFilter('oot',this)">OOT only</button>
        <button class="pill high-pill" onclick="setFilter('high',this)">High conf</button>
        <button class="pill low-pill" onclick="setFilter('low',this)">Manual req</button>
      </div>
      <div class="sidebar-label" style="margin-top:12px">Filter by axis</div>
      <div class="filter-row">
        <button class="pill" onclick="setAxisFilter('D',this)">Ø Diameter</button>
        <button class="pill" onclick="setAxisFilter('F',this)">GD&T</button>
        <button class="pill" onclick="setAxisFilter('M',this)">Distance</button>
      </div>
    </div>
    <div class="sidebar-section">
      <div class="sidebar-label">Files</div>
      <div style="display:flex;flex-direction:column;gap:5px">
        <div style="font-family:var(--mono);font-size:10px;display:flex;align-items:center;gap:6px">
          <span style="background:rgba(0,212,255,.15);color:var(--accent);padding:1px 5px;border-radius:2px;font-size:9px">CMM</span>
          {cmm_filename}
        </div>
        {''.join('<div style="font-family:var(--mono);font-size:10px;display:flex;align-items:center;gap:6px"><span style="background:rgba(0,200,122,.15);color:var(--ok);padding:1px 5px;border-radius:2px;font-size:9px">NC</span> ' + nc + '</div>' for nc in nc_filenames)}
      </div>
    </div>
  </aside>

  <main class="main">
    <div class="summary-bar">
      <div class="sum-item"><div class="sum-dot oot"></div><span style="color:var(--oot);font-weight:600">{len(oot_features)} OOT features</span></div>
      <div class="sum-item"><div class="sum-dot high"></div>{recs_high + recs_medium} auto-matched</div>
      <div class="sum-item"><div class="sum-dot acc"></div>{len(nc_holes)} NC hole positions</div>
      {'<div class="sum-item" style="margin-left:auto;color:var(--dim)">Max dev: <span style="color:var(--oot)">'+f'{max_dev:+.3f}mm</span> ({max_dev_feat.name if max_dev_feat else ""})</div>' if max_dev_feat else ''}
    </div>

    <div class="tabs">
      <div class="tab active" onclick="switchTab('recs',this)">OFFSET RECOMMENDATIONS</div>
      <div class="tab" onclick="switchTab('ops',this)">NC OPERATIONS ({len(ops_summary)})</div>
    </div>

    <div id="tab-recs" class="tab-content active">
      <div class="results">
        <div class="section-header">
          <span class="section-title">Recommendations</span>
          <span class="section-count" id="recCount">{len(recs)} features</span>
        </div>
        <div class="col-header-row">
          <div></div>
          <div class="col-header">Feature</div>
          <div class="col-header">Nominal</div>
          <div class="col-header">Measured</div>
          <div class="col-header">Deviation</div>
          <div class="col-header">Tol used</div>
          <div class="col-header">Tool / Match</div>
          <div class="col-header">Correction</div>
        </div>
        <div id="recList"></div>
        <div id="detailPanel"></div>
      </div>
    </div>

    <div id="tab-ops" class="tab-content">
      <div class="results">
        <table class="ops-table">
          <thead><tr>
            <th>OP</th><th>Description</th><th>Tool</th><th>Tool desc</th><th>Holes</th>
          </tr></thead>
          <tbody>{ops_rows}</tbody>
        </table>
      </div>
    </div>
  </main>
</div>

<script>
const RECS = {recs_json};
let state = {{filter:'all', axisFilter:null, selected:null}};

function getFiltered(){{
  let r = RECS;
  if(state.filter==='oot')  r = r.filter(x=>x.feature.out_of_tol);
  if(state.filter==='high') r = r.filter(x=>x.confidence==='HIGH');
  if(state.filter==='low')  r = r.filter(x=>x.confidence==='LOW');
  if(state.axisFilter)      r = r.filter(x=>x.feature.axis===state.axisFilter);
  return r;
}}

function render(){{
  const recs = getFiltered();
  document.getElementById('recCount').textContent = recs.length + ' features';
  document.getElementById('recList').innerHTML = recs.map((r,i) => rowHtml(r,i)).join('');
  document.getElementById('detailPanel').innerHTML = '';
  state.selected = null;
}}

function rowHtml(r,idx){{
  const f = r.feature;
  const pct = Math.min(f.tol_used_pct, 150);
  const barCls = pct>100?'oot':pct>80?'warn':'ok';
  const devCls = f.dev>=0?'pos':'neg';
  const corrCls = r.correction>0?'up':'down';
  const corrArrow = r.correction>0?'▲':'▼';
  const midTag = r.using_midpoint
    ? `<span style="font-family:var(--mono);font-size:9px;color:var(--warn);border:1px solid rgba(255,165,2,.3);padding:1px 4px;border-radius:2px;margin-left:4px">MID</span>`
    : '';
  const tolLabel = f.plus_tol===f.minus_tol
    ? `±${{f.plus_tol.toFixed(3)}}`
    : `+${{f.plus_tol.toFixed(3)}} / -${{f.minus_tol.toFixed(3)}}`;
  const toolHtml = r.tool==='—'
    ? `<button class="tool-assign-btn" onclick="event.stopPropagation();showToast('Manual assignment — coming soon')">+ Assign tool</button>`
    : `<div class="tool-info"><span class="tool-num">T${{r.tool}}</span><span class="tool-desc-text" title="${{r.tool_desc}}">${{r.tool_desc}}</span></div>`;
  return `<div class="rec-row ${{f.out_of_tol?'oot':'ok'}}" onclick="selectRec(${{idx}})">
    <div class="rec-stripe ${{f.out_of_tol?'oot':'ok'}}"></div>
    <div class="rec-cell">
      <span class="rec-feature-name">${{f.name}}</span>
      <span class="rec-feature-sub">${{f.type}} · AX:${{f.axis}} · pg${{f.page}}</span>
      <span style="margin-top:2px"><span class="status-badge ${{f.status}}">${{f.status}}</span></span>
    </div>
    <div class="rec-cell">
      <span class="val-main">${{f.nominal.toFixed(3)}}</span>
      <span class="val-sub">${{tolLabel}}</span>
    </div>
    <div class="rec-cell">
      <span class="val-main">${{f.meas.toFixed(3)}}</span>
      <span class="val-sub">mm</span>
    </div>
    <div class="rec-cell">
      <span class="deviation-val ${{devCls}}">${{f.dev>=0?'+':''}}${{f.dev.toFixed(4)}}</span>
      <span class="val-sub">mm</span>
    </div>
    <div class="rec-cell bar-container">
      <div class="tol-bar-wrap"><div class="tol-bar ${{barCls}}" style="width:${{Math.min(pct,100)}}%"></div></div>
      <span class="val-sub">${{f.tol_used_pct}}% of tol</span>
      <span style="margin-top:2px"><span class="conf-badge ${{r.confidence}}">${{r.confidence}}</span></span>
    </div>
    <div class="rec-cell">${{toolHtml}}</div>
    <div class="rec-cell">
      ${{f.out_of_tol || Math.abs(r.correction)>0.0001
        ? `<span class="correction-val ${{corrCls}}">${{corrArrow}} ${{Math.abs(r.correction).toFixed(4)}}${{midTag}}</span><span class="val-sub">${{r.direction.split(' ')[0]}}</span>`
        : `<span class="val-sub">—</span>`}}
    </div>
  </div>`;
}}

function selectRec(idx){{
  const recs = getFiltered();
  if(state.selected===idx){{ state.selected=null; document.getElementById('detailPanel').innerHTML=''; return; }}
  state.selected = idx;
  const r = recs[idx]; const f = r.feature;
  document.getElementById('detailPanel').innerHTML = `
    <div class="detail-panel">
      <div class="detail-header">
        <span class="detail-title">${{f.name}} — ${{f.type}} [${{f.axis}}]</span>
        <button class="detail-close" onclick="document.getElementById('detailPanel').innerHTML='';state.selected=null">✕</button>
      </div>
      <div class="detail-body">
        <div class="detail-group">
          <div class="detail-group-label">Measurement</div>
          <div class="detail-row"><span class="detail-key">Nominal</span><span class="detail-val">${{f.nominal.toFixed(4)}} mm</span></div>
          <div class="detail-row"><span class="detail-key">+Tolerance</span><span class="detail-val">+${{f.plus_tol.toFixed(4)}} mm</span></div>
          <div class="detail-row"><span class="detail-key">−Tolerance</span><span class="detail-val">−${{f.minus_tol.toFixed(4)}} mm</span></div>
          ${{f.tol_midpoint!==0
            ? `<div class="detail-row"><span class="detail-key" style="color:var(--warn)">Band centre</span><span class="detail-val" style="color:var(--warn)">${{f.tol_center.toFixed(4)}} mm</span></div>`
            : ''}}
          <div class="detail-row"><span class="detail-key">Measured</span><span class="detail-val">${{f.meas.toFixed(4)}} mm</span></div>
          <div class="detail-row"><span class="detail-key">Deviation</span>
            <span class="detail-val" style="color:${{f.dev>=0?'var(--ok)':'var(--oot)'}}">${{f.dev>=0?'+':''}}${{f.dev.toFixed(4)}} mm</span></div>
          <div class="detail-row"><span class="detail-key">Out of tol</span>
            <span class="detail-val" style="color:var(--oot)">${{f.outtol.toFixed(4)}} mm</span></div>
          <div class="detail-row"><span class="detail-key">Tol consumed</span><span class="detail-val">${{f.tol_used_pct}}%</span></div>
          <div class="detail-row"><span class="detail-key">CMM page</span><span class="detail-val">${{f.page}}</span></div>
        </div>
        <div class="detail-group">
          <div class="detail-group-label">Tool Match</div>
          <div class="detail-row"><span class="detail-key">Tool no.</span><span class="detail-val">T${{r.tool}}</span></div>
          <div class="detail-row"><span class="detail-key">Operation</span><span class="detail-val">OP${{r.op_num}}</span></div>
          <div class="detail-row"><span class="detail-key">Confidence</span><span class="detail-val"><span class="conf-badge ${{r.confidence}}">${{r.confidence}}</span></span></div>
          <div class="detail-row"><span class="detail-key">Matched holes</span><span class="detail-val">${{r.n_holes}}</span></div>
          <div style="margin-top:10px;font-size:11px;color:var(--dim);line-height:1.6">${{r.match_reason}}</div>
        </div>
        <div class="detail-group">
          <div class="detail-group-label">Offset Recommendation</div>
          <div class="detail-row"><span class="detail-key">Direction</span><span class="detail-val">${{r.direction}}</span></div>
          <div class="detail-row"><span class="detail-key">Target</span>
            <span class="detail-val" style="color:var(--accent)">${{r.correction_target.toFixed(4)}} mm
              ${{r.using_midpoint?'<span style="font-size:9px;color:var(--warn)">(band centre)</span>':''}}
            </span>
          </div>
          <div class="detail-row"><span class="detail-key">Correction</span>
            <span class="detail-val" style="color:${{r.correction>0?'var(--ok)':'var(--oot)'}};font-size:18px;font-weight:700">
              ${{r.correction>0?'▲':'▼'}} ${{Math.abs(r.correction).toFixed(4)}} mm
            </span>
          </div>
          ${{r.n_holes>0
            ? `<div class="offset-box">Apply wear offset on T${{r.tool}}<br>OP${{r.op_num}}: ${{r.op_desc}}<br>${{r.correction>0?'Increase':'Decrease'}} tool comp by ${{Math.abs(r.correction).toFixed(4)}}mm<br>${{r.using_midpoint?'(correcting to tolerance band centre)':''}}</div>`
            : `<div style="margin-top:10px;font-size:11px;color:var(--dim)">Assign tool manually to generate specific offset instruction.</div>`}}
        </div>
      </div>
    </div>`;
  document.getElementById('detailPanel').scrollIntoView({{behavior:'smooth',block:'nearest'}});
}}

function setFilter(f,el){{
  state.filter=f;
  document.querySelectorAll('.sidebar .filter-row:first-of-type .pill').forEach(p=>p.classList.remove('active'));
  el.classList.add('active'); render();
}}
function setAxisFilter(a,el){{
  if(state.axisFilter===a){{state.axisFilter=null;el.classList.remove('active');}}
  else{{state.axisFilter=a;document.querySelectorAll('.sidebar .filter-row:last-of-type .pill').forEach(p=>p.classList.remove('active'));el.classList.add('active');}}
  render();
}}
function switchTab(name,el){{
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(t=>t.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('tab-'+name).classList.add('active');
}}
function showToast(msg){{
  const t=document.createElement('div');t.className='toast';t.textContent=msg;
  document.body.appendChild(t);setTimeout(()=>t.remove(),3500);
}}

render();
</script>
</body>
</html>"""


# ─── FILE PICKER GUI ─────────────────────────────────────────────────────────

def pick_files_gui() -> tuple:
    """
    Open a simple Tkinter GUI to select the CMM PDF and NC files.
    Returns (cmm_path, nc_paths, part_name) or exits if cancelled.
    """
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
    except ImportError:
        print("[ERROR] tkinter not available — use command-line arguments instead.")
        sys.exit(1)

    root = tk.Tk()
    root.title("CNCSETTER — Select Files")
    root.resizable(False, False)

    # ── Styling ──────────────────────────────────────────────
    BG      = "#0d0f12"
    SURFACE = "#141720"
    BORDER  = "#2e3347"
    TEXT    = "#c8cdd8"
    DIM     = "#5a6075"
    ACCENT  = "#00d4ff"
    OK      = "#00c87a"
    OOT     = "#ff4757"

    root.configure(bg=BG)
    root.geometry("520x440")

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TFrame", background=BG)
    style.configure("TLabel", background=BG, foreground=TEXT,
                    font=("Consolas", 10))
    style.configure("TButton", background=SURFACE, foreground=ACCENT,
                    font=("Consolas", 10), borderwidth=1, focusthickness=0)
    style.map("TButton",
              background=[("active", BORDER)],
              foreground=[("active", "#ffffff")])
    style.configure("TEntry", fieldbackground=SURFACE, foreground=TEXT,
                    font=("Consolas", 10), insertcolor=ACCENT)
    style.configure("Header.TLabel", foreground=ACCENT,
                    font=("Consolas", 14, "bold"), background=BG)
    style.configure("Dim.TLabel", foreground=DIM,
                    font=("Consolas", 9), background=BG)
    style.configure("OK.TLabel", foreground=OK,
                    font=("Consolas", 9), background=BG)

    # ── State ────────────────────────────────────────────────
    cmm_var    = tk.StringVar()
    nc_var     = tk.StringVar(value="No files selected")
    part_var   = tk.StringVar()
    out_var    = tk.StringVar(value="Same folder as CMM PDF (default)")
    nc_paths_  = []
    result     = {"cmm": None, "nc": [], "part": "", "output": None}

    # ── Layout ───────────────────────────────────────────────
    pad = dict(padx=20, pady=6)

    ttk.Label(root, text="⟳  CNCSETTER", style="Header.TLabel").pack(
        anchor="w", padx=20, pady=(18, 4))
    ttk.Label(root, text="Select your CMM report and NC programs",
              style="Dim.TLabel").pack(anchor="w", padx=20, pady=(0, 14))

    # Separator
    tk.Frame(root, height=1, bg=BORDER).pack(fill="x", padx=20, pady=(0, 14))

    # CMM row
    cmm_frame = ttk.Frame(root)
    cmm_frame.pack(fill="x", **pad)
    ttk.Label(cmm_frame, text="CMM Report  ", width=14).pack(side="left")
    ttk.Entry(cmm_frame, textvariable=cmm_var, width=32,
              state="readonly").pack(side="left", padx=(0, 8))

    def browse_cmm():
        p = filedialog.askopenfilename(
            title="Select CMM Report PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if p:
            cmm_var.set(Path(p).name)
            result["cmm"] = p
            # Auto-fill part name from filename if empty
            if not part_var.get():
                part_var.set(Path(p).stem)
            update_run_btn()

    ttk.Button(cmm_frame, text="Browse…", command=browse_cmm,
               width=10).pack(side="left")

    # NC row
    nc_frame = ttk.Frame(root)
    nc_frame.pack(fill="x", **pad)
    ttk.Label(nc_frame, text="NC Programs  ", width=14).pack(side="left")
    ttk.Label(nc_frame, textvariable=nc_var, style="Dim.TLabel",
              width=32, anchor="w").pack(side="left", padx=(0, 8))

    def browse_nc():
        paths = filedialog.askopenfilenames(
            title="Select NC Program Files (select multiple with Ctrl+click)",
            filetypes=[
                ("NC files", "*.nc *.cnc *.mpf *.txt"),
                ("All files", "*.*"),
            ],
        )
        if paths:
            nc_paths_.clear()
            nc_paths_.extend(paths)
            result["nc"] = list(paths)
            names = [Path(p).name for p in paths]
            if len(names) <= 3:
                nc_var.set(", ".join(names))
            else:
                nc_var.set(f"{', '.join(names[:2])}  (+{len(names)-2} more)")
            update_run_btn()

    ttk.Button(nc_frame, text="Browse…", command=browse_nc,
               width=10).pack(side="left")

    # Part name row
    part_frame = ttk.Frame(root)
    part_frame.pack(fill="x", **pad)
    ttk.Label(part_frame, text="Part name    ", width=14).pack(side="left")
    ttk.Entry(part_frame, textvariable=part_var, width=32).pack(
        side="left", padx=(0, 8))
    ttk.Label(part_frame, text="(optional)", style="Dim.TLabel").pack(side="left")

    # Output file row
    out_frame = ttk.Frame(root)
    out_frame.pack(fill="x", **pad)
    ttk.Label(out_frame, text="Save report  ", width=14).pack(side="left")
    ttk.Label(out_frame, textvariable=out_var, style="Dim.TLabel",
              width=32, anchor="w").pack(side="left", padx=(0, 8))

    def browse_output():
        # Suggest a filename based on the CMM file if already chosen
        initial = Path(result["cmm"]).stem + "_report.html" if result["cmm"] else "report.html"
        initial_dir = str(Path(result["cmm"]).parent) if result["cmm"] else "/"
        p = filedialog.asksaveasfilename(
            title="Save Report As",
            initialfile=initial,
            initialdir=initial_dir,
            defaultextension=".html",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")],
        )
        if p:
            result["output"] = p
            out_var.set(Path(p).name + "  (" + str(Path(p).parent) + ")")

    ttk.Button(out_frame, text="Browse…", command=browse_output,
               width=10).pack(side="left")

    # Separator
    tk.Frame(root, height=1, bg=BORDER).pack(fill="x", padx=20, pady=(10, 0))

    # Status label
    status_var = tk.StringVar(value="Select CMM PDF and at least one NC file to continue.")
    status_lbl = ttk.Label(root, textvariable=status_var, style="Dim.TLabel")
    status_lbl.pack(anchor="w", padx=20, pady=(8, 0))

    # Buttons row
    btn_frame = ttk.Frame(root)
    btn_frame.pack(fill="x", padx=20, pady=(12, 20))

    def on_cancel():
        root.destroy()
        sys.exit(0)

    def on_run():
        result["part"] = part_var.get().strip()
        # If no explicit output chosen, default to same folder as CMM PDF
        if not result["output"] and result["cmm"]:
            result["output"] = str(
                Path(result["cmm"]).parent / (Path(result["cmm"]).stem + "_report.html")
            )
        root.destroy()

    run_btn = ttk.Button(btn_frame, text="▶  Generate Report",
                         command=on_run, width=22)
    run_btn.pack(side="right", padx=(8, 0))
    run_btn.state(["disabled"])

    ttk.Button(btn_frame, text="Cancel", command=on_cancel,
               width=10).pack(side="right")

    def update_run_btn():
        if result["cmm"] and result["nc"]:
            run_btn.state(["!disabled"])
            n = len(result["nc"])
            status_var.set(
                f"Ready — 1 CMM report, {n} NC file{'s' if n>1 else ''}"
            )
        else:
            run_btn.state(["disabled"])

    root.mainloop()

    if not result["cmm"] or not result["nc"]:
        sys.exit(0)

    return result["cmm"], result["nc"], result["part"], result["output"]


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="CNCSETTER — Generate offset recommendation report from CMM PDF + NC files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python generate_report.py                                     (opens file picker)
  python generate_report.py report.pdf OP10.nc OP20.nc OP40.nc
  python generate_report.py report.pdf OP*.nc --output report_4313.html
  python generate_report.py report.pdf OP10.nc --part "4313-9100-90" --verbose
        """
    )
    parser.add_argument("cmm_pdf", nargs="?", default=None,
                        help="CMM report PDF (Calypso / PC-DMIS) — omit to open file picker")
    parser.add_argument("nc_files", nargs="*",
                        help="NC program file(s) — omit to open file picker")
    parser.add_argument("--output", "-o", default=None,
                        help="Output HTML file (default: <cmm_name>_report.html)")
    parser.add_argument("--part", "-p", default=None,
                        help="Part name/number (auto-detected from PDF if omitted)")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Show OCR progress bar per page")
    args = parser.parse_args()

    # ── If no files given on CLI, open the GUI picker ────────
    if not args.cmm_pdf or not args.nc_files:
        cmm_path, nc_paths, part_gui, output_gui = pick_files_gui()
        part_name_default = part_gui or Path(cmm_path).stem
        output_path = args.output or output_gui or (Path(cmm_path).stem + "_report.html")
    else:
        cmm_path          = args.cmm_pdf
        nc_paths          = args.nc_files
        part_name_default = args.part or Path(cmm_path).stem
        output_path       = args.output or (Path(cmm_path).stem + "_report.html")

    # Validate inputs
    if not os.path.isfile(cmm_path):
        print(f"[ERROR] CMM file not found: {cmm_path}")
        sys.exit(1)
    for p in nc_paths:
        if not os.path.isfile(p):
            print(f"[ERROR] NC file not found: {p}")
            sys.exit(1)

    try:
        from tqdm import tqdm as _tqdm
        _has_tqdm = True
    except ImportError:
        _has_tqdm = False

    import time

    def step(label):
        """Print a numbered step header with timestamp."""
        step.n = getattr(step, "n", 0) + 1
        print(f"\n[{step.n}] {label}")

    def done(msg):
        print(f"    ✓ {msg}")

    print("=" * 56)
    print("  CNCSETTER — Report Generator")
    print("=" * 56)
    print(f"  CMM : {Path(cmm_path).name}")
    print(f"  NC  : {', '.join(Path(p).name for p in nc_paths)}")
    print(f"  Out : {output_path}")
    print("=" * 56)

    # ── 1. Parse NC files ────────────────────────────────────
    step("Parsing NC files")
    t0 = time.time()
    all_holes: List[NCHole] = []
    all_ops: List[dict] = []

    nc_iter = nc_paths
    if _has_tqdm:
        nc_iter = _tqdm(
            nc_paths,
            desc="  NC files",
            unit="file",
            bar_format="  {l_bar}{bar:25}{r_bar}",
            ncols=56,
        )

    for p in nc_iter:
        holes, ops = parse_nc_file(p)
        all_holes.extend(holes)
        all_ops.extend(ops)
        drill_ops = sum(1 for o in ops if o["n_holes"] > 0)
        if not _has_tqdm:
            print(f"    {Path(p).name}: {len(ops)} ops, {len(holes)} positions, "
                  f"{drill_ops} drilling/reaming ops")

    done(f"{len(all_holes)} hole positions across {len(all_ops)} operations  "
         f"({time.time()-t0:.1f}s)")

    # ── 2. Parse CMM PDF ─────────────────────────────────────
    step(f"Parsing CMM PDF via OCR  ({Path(cmm_path).name})")
    import fitz as _fitz
    n_pages = len(_fitz.open(cmm_path))
    est_min = n_pages * 5 // 60
    est_sec = n_pages * 5 % 60
    print(f"    {n_pages} pages  —  estimated {est_min}m {est_sec}s")

    t0 = time.time()
    cmm_features = parse_cmm_pdf(cmm_path, verbose=True)
    oot = [f for f in cmm_features if f.out_of_tol]

    done(f"{len(cmm_features)} features parsed  "
         f"({len(oot)} OOT, {len(cmm_features)-len(oot)} OK)  "
         f"({time.time()-t0:.0f}s)")

    # ── 3. Correlate ─────────────────────────────────────────
    step("Correlating features to tools")
    t0 = time.time()
    recs = correlate(cmm_features, all_holes)
    high = sum(1 for r in recs if r.confidence == "HIGH" and r.feature.out_of_tol)
    med  = sum(1 for r in recs if r.confidence == "MEDIUM" and r.feature.out_of_tol)
    low  = sum(1 for r in recs if r.confidence == "LOW" and r.feature.out_of_tol)

    done(f"{len(recs)} recommendations  —  "
         f"HIGH:{high}  MEDIUM:{med}  LOW(manual):{low}  "
         f"({time.time()-t0:.2f}s)")

    # ── 4. Generate HTML ─────────────────────────────────────
    step("Generating HTML report")
    t0 = time.time()
    part_name = part_name_default
    html = generate_html(
        recs=recs,
        cmm_features=cmm_features,
        nc_holes=all_holes,
        ops_summary=all_ops,
        part_name=part_name,
        cmm_filename=Path(cmm_path).name,
        nc_filenames=[Path(p).name for p in nc_paths],
    )

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = os.path.getsize(output_path) / 1024
    done(f"Saved  {output_path}  ({size_kb:.0f} KB)  ({time.time()-t0:.2f}s)")

    print(f"\n{'='*56}")
    print(f"  ✓ Done — open {output_path} in any browser")
    print(f"    Self-contained, no internet required.")
    print(f"{'='*56}\n")


if __name__ == "__main__":
    main()
