#!/usr/bin/env python3
"""
BOQ -> Primavera P6 cost loader.

Reads a BOQ (Excel) and a P6 schedule (XER), matches every BOQ line to schedule activities,
splits each line across the zones/floors where that activity exists, checks the totals, and writes:

  * <out>.xer            - the original schedule + a PROJCOST (expense) table. Import in P6 via
                           File > Import > Primavera PM (XER) > Update Existing Project.
  * <out>-p6-import.xlsx - the same expenses in P6's spreadsheet format (File > Import > Spreadsheet).
  * <out>-review.xlsx    - every match with its reason + confidence, flags, totals check.

Two-step use (so a person - or Claude - approves the matches before anything is loaded):

  py scripts/boq_to_p6.py match  boq.xlsx schedule.xer --out outputs/job     # writes job-mapping.csv
  (review / edit job-mapping.csv: activity, split)
  py scripts/boq_to_p6.py build  boq.xlsx schedule.xer --out outputs/job     # uses job-mapping.csv

`py scripts/boq_to_p6.py run ...` does both in one go (no manual review - for testing).
"""
import argparse, csv, difflib, functools, json, math, re, sys
from collections import defaultdict, OrderedDict
from difflib import SequenceMatcher
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    sys.exit("openpyxl is required: py -m pip install openpyxl")

ROOT = Path(__file__).resolve().parent.parent
SYNONYMS_FILE = ROOT / "config" / "synonyms.json"

# ------------------------------------------------------------------ text helpers
STOP = {"and", "the", "of", "for", "to", "with", "in", "on", "at", "a", "an", "all", "etc", "including",
        "works", "work", "supply", "installation", "install", "system", "shall", "be", "as", "other",
        "than", "areas", "area", "no", "nos", "m", "m2", "m3", "mm", "per", "is", "by", "from", "up"}

def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def tokens(s):
    s = re.sub(r"\((?:excl|except|not incl)[^)]*\)", " ", str(s or ""), flags=re.I)  # "(excluding light fixtures)"
    s = re.sub(r"\ba\s*/\s*c\b", " hvac ", s, flags=re.I)
    words = re.findall(r"[a-z]+", s.lower().replace("/", " "))
    out = []
    for w in words:
        if w in STOP or len(w) < 2:
            continue
        if w.endswith("ies") and len(w) > 4:
            w = w[:-3] + "y"
        elif re.search(r"(sh|ch|x|ss|z)es$", w) and len(w) > 4:
            w = w[:-2]  # finishes -> finish, switches -> switch, boxes -> box
        elif w.endswith("s") and not w.endswith("ss") and len(w) > 3:
            w = w[:-1]
        out.append(w)
    return out

def load_synonyms():
    if SYNONYMS_FILE.exists():
        return json.loads(SYNONYMS_FILE.read_text(encoding="utf-8"))
    return {}

def expand(toks, syn):
    """Add canonical trade words for any synonym phrase found."""
    text = " " + " ".join(toks) + " "
    extra = []
    for canon, phrases in syn.items():
        for p in phrases:
            pt = tokens(p)
            if pt and " " + " ".join(pt) + " " in text:
                extra.extend(tokens(canon))
                break
    return toks + extra

def fuzzy_overlap(a, b):
    """Token overlap tolerant of misspellings (Coloumns/Columns, Pannel/Panel, Masonary/Masonry)."""
    if not a or not b:
        return 0.0
    return _overlap(frozenset(a), frozenset(b))

@functools.lru_cache(maxsize=500000)
def _overlap(a, bs):
    # cached: a big BOQ scores the same word sets against the same activity names over and over.
    # Cleared with WEIGHTS whenever a new schedule is loaded.
    hit = 0.0
    w = WEIGHTS.get
    for t in set(a):
        if t in bs:
            hit += w(t, 1.0)
        elif any(_stem_match(t, u) for u in bs) or (
                len(t) >= 5 and any(_similar(t, u) for u in bs if len(u) >= 5 and u[0] == t[0])):
            hit += 0.9 * w(t, 1.0)  # a spelling variant counts, but an exact word always wins a tie
    return hit / sum(w(t, 1.0) for t in set(a))

STEM_SUFFIXES = {"work", "works", "ing", "ings", "s", "es", "ed", "er", "ers", "ation", "ations", "age"}

@functools.lru_cache(maxsize=None)
def _stem_match(t, u):
    # duct/ductwork, pipe/pipework, finish/finishing, cable/cabling - real word endings only,
    # so 'base' never meets 'basement'
    a, b = (t, u) if len(t) <= len(u) else (u, t)
    if len(a) < 4 or a == b:
        return False
    for stem in {a, a[:-1] if a.endswith("e") else a}:
        if b.startswith(stem) and b[len(stem):] in STEM_SUFFIXES:
            return True
    return False

@functools.lru_cache(maxsize=None)
def _similar(t, u):
    # cached: a 1,300-line BOQ against ~700 activities repeats the same word pairs millions of times
    # misspellings (coloumns, pannel, masonary) - but not different words (chiller/chilled = 0.857)
    return difflib.SequenceMatcher(None, t, u).ratio() >= 0.88

# ------------------------------------------------------------------ XER
class Xer:
    def __init__(self, path):
        self.path = Path(path)
        raw = self.path.read_bytes()
        self.raw = raw
        # P6 desktop exports are cp1252 (checked: P6 23 + P6 25). Some tools write UTF-8. latin-1 is
        # the last resort because it round-trips any byte, so the original is never corrupted.
        for enc in ("utf-8", "cp1252", "latin-1"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        self.encoding = enc
        self.lines = text.splitlines()
        self.tables = OrderedDict(); self.fields = {}
        cur = None
        for ln in self.lines:
            p = ln.split("\t")
            if p[0] == "%T":
                cur = p[1]; self.tables[cur] = []
            elif p[0] == "%F":
                self.fields[cur] = p[1:]
            elif p[0] == "%R":
                self.tables[cur].append(dict(zip(self.fields[cur], p[1:])))
        if "TASK" not in self.tables:
            sys.exit(f"{path}: no TASK table - is this a P6 XER export?")
        projs = self.tables.get("PROJECT", [])
        if len(projs) != 1:
            names = [p.get("proj_short_name") for p in projs]
            sys.exit(f"{path}: contains {len(projs)} projects {names}. Export ONE project to XER and retry.")
        self.project = projs[0]
        self.wbs = {w["wbs_id"]: w for w in self.tables.get("PROJWBS", [])}

    def wbs_path(self, wbs_id):
        names = []
        while wbs_id in self.wbs:
            w = self.wbs[wbs_id]
            if w.get("proj_node_flag") == "Y":
                break
            names.append(w["wbs_name"])
            wbs_id = w.get("parent_wbs_id")
        return list(reversed(names))

    def activities(self):
        acts = []
        for t in self.tables["TASK"]:
            if t.get("task_type") in ("TT_Mile", "TT_FinMile", "TT_LOE", "TT_WBS"):
                continue  # milestones / LOE / WBS summary never carry BOQ cost
            acts.append({"task_id": t["task_id"], "code": t["task_code"], "name": t["task_name"],
                         "wbs": self.wbs_path(t["wbs_id"]), "status": t.get("status_code", "")})
        return acts

ZONE_WORDS = re.compile(r"\b(floor|level|lvl|basement|underground|roof|podium|zone|block|building|"
                        r"phase|area|sector|tower|wing|storey|story|ground|first|second|third|fourth|"
                        r"fifth|mezzanine|parking)\b", re.I)

def group_activities(acts):
    """Group activities that are the same work repeated per zone (e.g. 'Slabs-GF', 'Slabs-FF').
    Base name = activity name minus a trailing zone suffix; zone = the WBS node that looks like a
    floor/zone, else the suffix."""
    groups = defaultdict(list)
    for a in acts:
        name = a["name"]
        m = re.match(r"^(.*?)\s*[-_]\s*([A-Z0-9]{1,4})$", name)
        base, suffix = (m.group(1), m.group(2)) if m else (name, "")
        zones = [w for w in a["wbs"] if ZONE_WORDS.search(w)]
        zone = (zones[-1] if zones else "") or suffix  # deepest floor/zone node
        a["base"], a["zone"] = norm(base), zone
        groups[a["base"].lower()].append(a)
    return groups

# ------------------------------------------------------------------ BOQ
HDR = {
    "desc": re.compile(r"desc|scope of work|particulars|specification", re.I),
    "title": re.compile(r"main title|heading|title|section|trade|element|bill", re.I),
    "unit": re.compile(r"^\s*(unit|uom|u\.o\.m)\s*$", re.I),
    "qty": re.compile(r"^\s*(qty|quantity|quantities|qnty)\b", re.I),
    "rate": re.compile(r"^\s*(rate|unit\s*(rate|price|cost)|price)\b", re.I),
    "amount": re.compile(r"amount|total\s*(cost|price)?|value", re.I),
    "ref": re.compile(r"^\s*(sn|s\.n|no\.?|item|item\s*no|ref|#|sr)\s*$", re.I),
}
# columns a planner adds by hand - never read them as input (we must not peek at the answer)
IGNORE = re.compile(r"activity|type of work|building|floor|check|wbs|%", re.I)

TOTAL_ROW = re.compile(r"^(grand\s+)?total|^sub[\s-]*total|carried|brought forward|to collection|"
                       r"^collection", re.I)
# rows that are page furniture, not work headings
BANNER = re.compile(r"^(division\b|div\.?\s*\d|bill of quantities|summary\b|ref\.?\s*drawings?|series numbers|"
                    r"specification sections?|spec\.? sections?|notes?\s*[:\-]*$|the contractor|general notes)", re.I)
SUBITEM = re.compile(r"^\(?\d+(\.\d+)*\s*[.)]+\s")  # "1.) Access to site" - a sub-point, not a heading
FIRST_CLAUSE = re.compile(r"\s+(related to|including|includ\w*|comprising|as per|according to|complete with)\b"
                          r"|\s*[,:;(]", re.I)

ZONE_HEAD = re.compile(r"^\s*(lower\s*level|basement|level|lvl|ground|mezzanine|roof|podium|"
                       r"\d+\s*(st|nd|rd|th)(\s+mezzanine)?)[\s\-]*\d*\s*(floor|level)?\s*$", re.I)

def zone_key(s):
    """Normalise a floor/zone name so a BOQ heading ('4TH FLOOR', 'LOWER LEVEL-1') meets a schedule
    zone ('Level 04', 'Lower Level 1', 'F04', 'B1'). None if it is not a recognisable floor."""
    t = str(s or "").lower()
    n = re.search(r"(\d+)", t)
    if "mezz" in t or t.strip() in ("mz", "mez"):
        return "MZ"
    if "ground" in t or t.strip() in ("gf", "g"):
        return "GF"
    if "roof" in t or t.strip() == "rf":
        return "RF"
    if re.search(r"lower|basement|^b\d|^ll\d|under", t):
        return f"B{int(n.group(1))}" if n else "B"
    if n and re.search(r"floor|level|lvl|^f\d|^l\d|\d(st|nd|rd|th)", t):
        return f"F{int(n.group(1))}"
    return None

def _is_caps(t):
    letters = [c for c in t if c.isalpha()]
    return bool(letters) and sum(c.isupper() for c in letters) / len(letters) >= 0.8

def _looks_like_codes(t):
    letters = sum(c.isalpha() for c in t)
    return letters < 0.3 * max(1, len(t.replace(" ", "")))

BULLET = re.compile(r"^[­\-–•*·]")                     # "­ Maximum aggregate size ..." = a note
ATTRIBUTE = re.compile(r"^[^:\-–]{2,35}?\s+[-–:]\s+\S")        # "Speed   -   1 m/s" = detail of the item above
LEAD_VERBS = re.compile(r"^(?:(?:supply(?:ing)?|install(?:ation|ing)?|testing|test|commissioning|provide|"
                        r"providing|furnish(?:ing)?|fix(?:ing)?|apply(?:ing)?|erect(?:ion|ing)?|and|&|,)\s*)+(?:of\s+|the\s+)*", re.I)

def pick_heading(block):
    """From the text rows between two priced lines, pick the work heading the way a planner does:
    the last ALL-CAPS short row ('INTERNAL FINISHES' > 'BASE FINISHES' -> BASE FINISHES), else the
    last short row ('Precast concrete'), else the first clause of the spec paragraph
    ('Reinforced poured concrete related to ...' -> 'Reinforced poured concrete';
    'Supply, installation ... of elevator equipment in ...' -> 'elevator equipment in ...')."""
    cand = [t for t in block if not SUBITEM.match(t) and not BULLET.match(t) and not ATTRIBUTE.match(t)
            and not re.search(r"[:：]\s*-?\s*$", t)]  # "Reinforcing steel ... (Fy) :-" introduces notes
    short = [t for t in cand if len(t) <= 60]
    caps = [t for t in short if _is_caps(t)]
    if caps:
        return caps[-1]
    if short:
        return short[-1]
    for t in cand:
        head = FIRST_CLAUSE.split(LEAD_VERBS.sub("", t), maxsplit=1)[0].strip()[:60]
        if len(tokens(head)) >= 1 and len(head) >= 3:
            return head
    return ""

def find_header(rows):
    for i, r in enumerate(rows[:40]):
        found = {}
        for j, c in enumerate(r):
            if c is None or IGNORE.search(str(c)):
                continue
            for k, rx in HDR.items():
                if k not in found and rx.search(str(c)):
                    found[k] = j
                    break
        if "desc" in found and ("qty" in found or "amount" in found):
            return i, found
    return None, {}

def xlsx_sheets(path, only=None):
    """Cell VALUES of every sheet, as [(name, rows)] with row 1 at index 0 (empty rows kept, so row
    numbers match Excel). Reads the xlsx XML directly: real tender BOQs often carry 100,000+
    bloated cell styles that make openpyxl take minutes. Falls back to openpyxl if anything is odd."""
    import zipfile, xml.etree.ElementTree as ET
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    try:
        z = zipfile.ZipFile(path)
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(NS + "si"):
                shared.append("".join(t.text or "" for t in si.iter(NS + "t")))
        rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        out = []
        for sh in ET.fromstring(z.read("xl/workbook.xml")).iter(NS + "sheet"):
            name = sh.get("name")
            if only and name != only:
                continue
            target = rels[sh.get(RNS + "id")].lstrip("/")
            target = target if target.startswith("xl/") else "xl/" + target
            grid = {}
            for row in ET.fromstring(z.read(target)).iter(NS + "row"):
                r = int(row.get("r"))
                for c in row.iter(NS + "c"):
                    ref, t = c.get("r"), c.get("t")
                    col = 0
                    for ch in ref:
                        if ch.isalpha():
                            col = col * 26 + ord(ch.upper()) - 64
                    v = c.find(NS + "v")
                    if t == "s" and v is not None:
                        val = shared[int(v.text)]
                    elif t == "inlineStr":
                        val = "".join(x.text or "" for x in c.iter(NS + "t"))
                    elif t in ("str", "e"):
                        val = v.text if (v is not None and t == "str") else None
                    elif t == "b":
                        val = bool(int(v.text)) if v is not None else None
                    elif v is not None and v.text is not None:
                        num = float(v.text)
                        val = int(num) if num.is_integer() else num
                    else:
                        val = None
                    if val not in (None, ""):
                        grid.setdefault(r, {})[col - 1] = val
            n_rows = max(grid, default=0)
            width = max((max(cols) + 1 for cols in grid.values()), default=0)
            out.append((name, [tuple(grid.get(i, {}).get(j) for j in range(width))
                               for i in range(1, n_rows + 1)]))
        return out
    except Exception:
        wb = load_workbook(path, data_only=True, read_only=True)
        return [(ws.title, list(ws.iter_rows(values_only=True)))
                for ws in ([wb[only]] if only else wb.worksheets)]

def boq_files(path):
    p = Path(path)
    if p.is_dir():
        files = sorted(f for f in p.glob("*.xls*") if not f.name.startswith("~$"))
        if not files:
            sys.exit(f"{path}: no Excel files in this folder.")
        return files
    return [p]

def read_boq(path, sheet=None):
    """Read one BOQ workbook, or a folder of them (e.g. one file per CSI division).
    Works with a heading column (planner layout) or with headings as text rows above the priced
    lines (typical tender BOQ) - then the heading is carried down to every line under it."""
    files = boq_files(path)
    multi = len(files) > 1
    items, header = [], {}
    for f in files:
        # first sheet with a BOQ header (later sheets are usually the planner's own working/pivots)
        for ws_name, rows in xlsx_sheets(f, sheet):
            hdr_i, cols = find_header(rows)
            if hdr_i is None:
                continue  # notes / scratch sheet
            header = header or {k: rows[hdr_i][v] for k, v in cols.items()}
            division, heading, block, context, open_block = "", "", [], "", False
            zone_hint = ""
            floor_heading_pending = False
            for n, r in enumerate(rows[hdr_i + 1:], start=hdr_i + 2):
                g = lambda k: r[cols[k]] if k in cols and cols[k] < len(r) else None
                desc, title = norm(g("desc")), norm(g("title"))
                num = lambda v: float(v) if isinstance(v, (int, float)) else _num(v)
                qty, rate, amt = num(g("qty")), num(g("rate")), num(g("amount"))
                label = desc or title or norm(g("ref"))
                if not label:
                    implied = amt / rate if amt and rate else 0
                    # a real lumped price has a rate AND an amount that is a whole number of units
                    # (535,000 = 2 x 267,500); a bare amount is a collection-page subtotal - skip it
                    if implied >= 1 and abs(implied - round(implied)) < 1e-6:
                        key = f"{f.stem}:{n}" if multi else n
                        items.append({"row": key, "ref": f"{f.stem}-{n}" if multi else str(n),
                                      "title": heading, "desc": "", "context": context,
                                      "division": division or f.stem, "file": f.name, "unit": "",
                                      "qty": None, "rate": rate, "amount": amt,
                                      "layout_flag": "Price row with no description or quantity - the BOQ "
                                                     "lumps several items into one amount. Check which "
                                                     "items it covers and split it in the mapping file."})
                        open_block = False
                    continue
                if qty is None and TOTAL_ROW.match(label):
                    open_block = False  # subtotal ("To Collection") - never a BOQ line
                    continue
                if amt is None and qty is not None and rate is not None:
                    amt = qty * rate
                if qty is None or amt is None:  # a text row: heading, spec paragraph or banner
                    if re.match(r"^(division|div\.?)\s*\d", desc, re.I):
                        division = desc
                    if not desc or BANNER.match(desc) or _looks_like_codes(desc):
                        continue
                    if ZONE_HEAD.match(desc) and zone_key(desc):
                        zone_hint = desc  # "4TH FLOOR" = where the next lines go, not what they are
                        floor_heading_pending = True
                        continue
                    if not open_block:
                        block, open_block = [], True
                    block.append(desc)
                    if len(desc) > 60:
                        context = desc
                    h = pick_heading(block)
                    if h and h != heading and not floor_heading_pending:
                        zone_hint = ""  # a new work heading ends the previous floor's lines
                    heading = h or heading
                    continue
                open_block = False
                floor_heading_pending = False
                key = f"{f.stem}:{n}" if multi else n
                items.append({"row": key, "ref": (f"{f.stem}-{n}" if multi else norm(g("ref")) or str(n)),
                              "title": title or heading, "desc": desc, "context": context,
                              "heading_src": "column" if title else "carried", "zone_hint": zone_hint,
                              "division": division or f.stem, "file": f.name,
                              "unit": norm(g("unit")), "qty": qty, "rate": rate, "amount": amt})
            break
    if not items:
        sys.exit(f"{path}: no priced BOQ lines found (need Description + Qty + Rate/Amount columns).")
    return items, header

def _num(v):
    if v is None:
        return None
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return float(s) if s not in ("", "-", ".") else None
    except ValueError:
        return None

# ------------------------------------------------------------------ matching
# word weights: a word used in many activity names ('finishes', 'water') decides less than a word
# used in one ('ceiling'). Set per run from the schedule's own activity names (IDF-style).
WEIGHTS = {}

# words every BOQ description repeats ("supply, installation and testing of ...") - they say nothing
# about WHICH work it is, so they must not pull lines onto a 'Testing & Commissioning' activity
BOILERPLATE = {"supply", "installation", "install", "installed", "testing", "test", "tested", "commissioning",
               "commission", "provide", "providing", "fixing", "fix", "complete", "necessary", "accessory",
               "ancillary", "required", "requirement", "rate", "item", "drawing", "specification", "section",
               "include", "includes", "including", "according", "approved", "contractor", "etc", "type"}

@functools.lru_cache(maxsize=None)
def _known(t):
    """Is this word (or its stem: duct -> ductwork) used in any activity name of this schedule?
    Cached; the cache is cleared whenever WEIGHTS is reset for a new schedule."""
    return t in WEIGHTS or any(_stem_match(t, u) for u in WEIGHTS)

def boq_words(s):
    return [t for t in tokens(s) if t not in BOILERPLATE]

def score(text_toks, group_toks, base=None):
    """How well a BOQ text fits an activity name: share of the activity's words found in the text
    (with and without synonym expansion) and, for a short heading, share of the heading's words
    found in the activity name ('Metal Doors' -> 'Internal Doors & Ironmongery')."""
    s = fuzzy_overlap(group_toks, text_toks)
    if base:
        s = max(s, fuzzy_overlap(base, text_toks))
    # short text: share of its words found in the activity name - counting only words the schedule
    # uses at all ('reinforced', 'thick' appear in no activity name, so they neither help nor hurt)
    known = [t for t in set(text_toks) if _known(t)]
    if known and len(set(text_toks)) <= 6:
        # one known word ('PIPES') can decide, a little less surely than two
        s = max(s, 0.95 * (1.0 if len(known) >= 2 else 0.85) * fuzzy_overlap(known, group_toks))
    return s

CSI_FILE = ROOT / "config" / "csi_divisions.json"

def load_csi():
    if not CSI_FILE.exists():
        return {}
    data = json.loads(CSI_FILE.read_text(encoding="utf-8"))
    return {int(k): [tokens(w) for w in v] for k, v in data.items() if k.isdigit()}

def div_number(s):
    m = re.search(r"\bdiv(?:ision)?\.?\s*0*(\d{1,2})\b", str(s or ""), re.I)
    return int(m.group(1)) if m else None

def affinity(gt, words):
    """1 if the activity's name uses any of the division's trade words (phrase-aware), else 0."""
    text = " " + " ".join(gt) + " "
    return 1.0 if any(w and (" " + " ".join(w) + " ") in text for w in words) else 0.0

def match_items(items, groups, syn, csi=None):
    csi = csi or {}
    gbase = {k: tokens(v[0]["base"]) for k, v in groups.items()}
    # activity names are used as written - synonyms only translate BOQ wording. (Expanding the names
    # too gave 'Hose Cabinets' the word 'casework' and pulled joinery onto a fire-fighting activity.)
    gtoks = dict(gbase)
    df = defaultdict(int)
    for g in gtoks.values():
        for t in set(g):
            df[t] += 1
    WEIGHTS.clear(); _known.cache_clear(); _overlap.cache_clear()
    WEIGHTS.update({t: 1 / math.sqrt(n) for t, n in df.items()})
    results = []
    for it in items:
        tt = expand(tokens(it["title"]), syn)
        dt = expand(boq_words(it["desc"]), syn)
        dv = div_number(it.get("division")) or div_number(it.get("file"))
        dwords = csi.get(dv, [])
        aff = {k: affinity(g, dwords) for k, g in gtoks.items()} if dwords else {}
        # the division is context: it breaks ties and demotes cross-trade matches (a Div 26 line
        # should never land on a doors activity just because the words overlap)
        def best(text):
            if not text:
                return (0, None)
            return max(((score(text, g, gbase[k]) * (1.0 if not aff or aff[k] else 0.6) + 0.05 * aff.get(k, 0), k)
                        for k, g in gtoks.items()), default=(0, None))
        ct = expand(boq_words(it.get("context", "")), syn)
        best_t, best_d, both = best(tt), best(dt), best(tt + dt)
        if best_t[1] and dt:
            # a generic heading ('reinforced concrete') fits several activities equally - let the
            # line's own description decide between them (slab / column / raft / wall)
            sc = {k: score(tt, g, gbase[k]) * (1.0 if not aff or aff[k] else 0.6) + 0.05 * aff.get(k, 0)
                  for k, g in gtoks.items()}
            tied = [k for k, v in sc.items() if v >= best_t[0] - 0.05]
            if len(tied) > 1:
                pick = max(tied, key=lambda k: (score(dt, gtoks[k], gbase[k]), sc[k]))
                if score(dt, gtoks[pick], gbase[pick]) > 0:
                    best_t = (best_t[0], pick)  # tied on the heading - same heading score
        best_c = best(ct) if it.get("heading_src") == "carried" else (0, None)
        flags = [it["layout_flag"]] if it.get("layout_flag") else []
        # heading is the planner's primary key; description decides only if there is no heading
        if it["title"] and best_t[0] >= 0.5:
            s, key, basis = best_t[0], best_t[1], "heading"
        elif best_c[0] >= 0.5:  # weak heading ("2140 L/S Outdoor Air"): the spec paragraph above names it
            s, key, basis = best_c[0], best_c[1], "spec paragraph"
        elif best_d[0] >= 0.5:
            s, key, basis = best_d[0], best_d[1], "description"
        else:
            s, key, basis = both[0], both[1], "heading+description"
        s = min(s, 1.0)
        # description cross-checks only make sense when a person typed the heading into a column;
        # a heading carried down from the BOQ layout came from that same text
        if it.get("heading_src") == "column" and key:
            key_in_desc = score(dt, gtoks[key])
            if best_d[0] >= 0.85 and best_d[1] != key and key_in_desc < 0.5:
                flags.append(f"Description reads like '{groups[best_d[1]][0]['base']}', "
                             f"but heading says '{it['title']}'. Check which one is right.")
            elif key_in_desc < 0.5 and len(tokens(it["desc"])) >= 5:
                flags.append("Description does not mention this activity's work - check the line is filed "
                             "under the right heading.")
        if dwords and key and not aff.get(key):
            flags.append(f"Matched activity is outside this line's trade (Division {dv}) - check it.")
        if key and s >= 0.5:
            src = {"heading": tt, "spec paragraph": ct, "description": dt}.get(basis, tt + dt)
            shared = {t for t in set(src) if any(t == u or _stem_match(t, u) or _similar(t, u)
                                                  for u in gtoks[key])}
            if len(shared) == 1 and len(set(gtoks[key])) > 1:  # 'Excavation' = 'Excavation' is a full match
                note = f"matched on one word ('{next(iter(shared))}')"
            else:
                note = ""
        else:
            note = ""
        s_raw = s - 0.05 * aff.get(key, 0) if key else s  # the division bonus breaks ties, never earns HIGH
        if s_raw >= 0.85 and not flags and not note:
            conf = "HIGH"
        elif s >= 0.5:
            conf = "MEDIUM"
        else:
            guess = groups[key][0]["base"] if key and s > 0 else None
            conf, key = "LOW", None  # a weak guess never loads money - wrong money is worse than none
            flags.append("No clear activity for this line - assign it in the mapping file"
                         + (f" (weak guess: '{guess}')." if guess else "."))
        results.append({**it, "group": key, "score": round(s, 2), "conf": conf, "flags": flags,
                        "basis": basis + (f" - {note}" if note else "")})
    return results

def rate_outliers(results):
    """Same heading + same unit but rates more than 5x apart = likely a typo."""
    by = defaultdict(list)
    for r in results:
        if r["rate"]:
            by[(r["title"].lower(), r["unit"].lower())].append(r)
    for rows in by.values():
        if len(rows) < 2:
            continue
        rates = sorted(x["rate"] for x in rows)
        if len(rows) == 2:  # with two lines we can't tell which one is wrong - flag both
            if rates[0] and rates[1] / rates[0] > 5:
                for x in rows:
                    other = next(o for o in rows if o is not x)
                    x["flags"].append(f"Rate {x['rate']:,.2f}/{x['unit']} vs {other['rate']:,.2f} on line "
                                      f"{other['ref']} (same heading) - {rates[1] / rates[0]:.0f}x apart, "
                                      "one is likely a typo.")
            continue
        med = rates[len(rates) // 2]
        for x in rows:
            if med and (x["rate"] / med > 5 or med / x["rate"] > 5):
                x["flags"].append(f"Rate {x['rate']:,.2f}/{x['unit']} is {max(x['rate'] / med, med / x['rate']):.0f}x "
                                  f"the typical {med:,.2f} for this heading - possible typo.")

def duplicate_lines(results):
    seen = defaultdict(list)
    for r in results:
        seen[(r["desc"].lower(), round(r["amount"], 2))].append(r)
    for rows in seen.values():
        if len(rows) > 1 and rows[0]["desc"]:
            refs = ", ".join(x["ref"] for x in rows)
            for x in rows:
                x["flags"].append(f"Same description and amount as lines {refs} - possible duplicate.")

# ------------------------------------------------------------------ mapping file (the human checkpoint)
MAP_COLS = ["ref", "row", "heading", "description", "amount", "activity_group", "split", "confidence",
            "basis", "flags"]

def write_mapping(results, groups, path):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(MAP_COLS)
        for r in results:
            g = groups.get(r["group"]) if r["group"] else None
            w.writerow([r["ref"], r["row"], r["title"], r["desc"][:200], round(r["amount"], 2),
                        g[0]["base"] if g else "", "equal", r["conf"], r["basis"], " | ".join(r["flags"])])

def read_mapping(path, groups):
    """activity_group = base activity name (or a single Activity ID); split = 'equal' or
    'ZONE:weight;ZONE:weight' (e.g. 'GF:50;FF:30;SF:20')."""
    by_base = {g[0]["base"].lower(): k for k, g in groups.items()}
    by_code = {a["code"].lower(): a for g in groups.values() for a in g}
    out = {}
    with open(path, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            target = (row.get("activity_group") or "").strip()
            key = by_base.get(target.lower())
            single = by_code.get(target.lower())
            if target and not key and not single:
                sys.exit(f"{path}: line {row['ref']}: activity '{target}' is not in the schedule.")
            if str(row["row"]) in out:
                sys.exit(f"{path}: duplicate BOQ row {row['row']} in mapping.")
            out[str(row["row"])] = {"group": key, "single": single,
                                    "split": (row.get("split") or "equal").strip(),
                                    "ref": row.get("ref"), "amount": row.get("amount")}
    return out

# ------------------------------------------------------------------ BOQ revision (keep approved decisions)
REV_COLS = ["revision_status", "previous_amount", "previous_activity", "previous_split", "revision_note"]
REVIEW_STATUS = ("NEW", "CHANGED", "REWORDED", "SCHEDULE_CHANGED", "UNASSIGNED")

def line_id(heading, desc):
    """Row numbers shift when a line is inserted, so a BOQ line is known by its heading + description."""
    return norm(heading).lower(), norm((desc or "")[:200]).lower()

def read_previous(path):
    with open(path, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    if not rows or not {"heading", "description", "amount", "activity_group", "split"} <= set(rows[0]):
        sys.exit(f"{path}: not a mapping file from this tool.")
    for x in rows:
        x["id"] = line_id(x["heading"], x["description"])
        x["amt"] = _num(x["amount"]) or 0.0
    return rows

def diff_boq(prev, results):
    """Pair revised lines with previous ones. Returns {id(result): (old, kind)} and the deleted lines.
    Identical lines pair on equal amounts first, so a repeated item keeps its own decision."""
    olds = defaultdict(list)
    for x in prev:
        olds[x["id"]].append(x)
    pairs = {}
    ident = lambda r: line_id(r["title"], r["desc"])
    for exact in (True, False):
        for r in results:
            pool = olds.get(ident(r))
            if id(r) in pairs or not pool:
                continue
            hit = next((x for x in pool if abs(x["amt"] - r["amount"]) <= 0.01), None if exact else pool[0])
            if hit:
                pool.remove(hit)
                pairs[id(r)] = (hit, "same" if abs(hit["amt"] - r["amount"]) <= 0.01 else "changed")
    left = [x for pool in olds.values() for x in pool]
    # not an exact match: same description under another heading, or small rewording under the same one.
    # The scope may have moved, so these only ever suggest the old decision
    for r in results:
        if id(r) in pairs or not r["desc"]:
            continue
        h, d = ident(r)
        best, best_s = None, 0.85
        for x in left:
            s = 1.0 if x["id"][1] == d else (
                SequenceMatcher(None, d, x["id"][1]).ratio() if x["id"][0] == h and x["id"][1] else 0)
            if s > best_s or (s == best_s and not best):
                best, best_s = x, s
        if best:
            left.remove(best)
            pairs[id(r)] = (best, "reworded")
    return pairs, left

def split_problem(split, targets):
    """A custom split naming a zone or activity the revised schedule no longer has."""
    if not split or split.lower() == "equal":
        return None
    have = {a["zone"].lower() for a in targets} | {a["code"].lower() for a in targets}
    gone = [p.split(":")[0].strip() for p in split.split(";")
            if ":" in p and p.split(":")[0].strip().lower() not in have]
    return ", ".join(gone) or None

def revise(results, groups, prev):
    """Carry approved decisions onto the revised BOQ; every line gets a revision_status."""
    by_base = {g[0]["base"].lower(): k for k, g in groups.items()}
    by_code = {a["code"].lower(): a for g in groups.values() for a in g}
    pairs, deleted = diff_boq(prev, results)
    for r in results:
        fresh = groups[r["group"]][0]["base"] if r["group"] else ""
        old, kind = pairs.get(id(r), (None, None))
        r.update(map_activity=fresh, map_split="equal", prev_amt="", prev_act="", prev_split="", rev_note="")
        if not old:
            r["rev_status"], r["rev_note"] = "NEW", "New line - matched fresh."
            continue
        act, split = old["activity_group"].strip(), (old["split"] or "equal").strip()
        r.update(prev_amt=round(old["amt"], 2), prev_act=act, prev_split=split)
        if kind == "reworded":
            r["rev_status"] = "REWORDED"
            r["rev_note"] = (f"Was '{old['description'][:80]}' under '{old['heading']}' - matched fresh, "
                             "check the previous decision still applies.")
            continue
        key, single = by_base.get(act.lower()), by_code.get(act.lower())
        if act and not key and not single:
            r["rev_status"] = "SCHEDULE_CHANGED"
            r["rev_note"] = f"Activity '{act}' is no longer in the schedule - matched fresh."
            continue
        gone = split_problem(split, [single] if single else (groups.get(key) or []))
        if gone:
            r.update(map_activity=act, map_split="equal", rev_status="SCHEDULE_CHANGED",
                     rev_note=f"Split '{split}' names {gone}, no longer on this activity - reset to equal.")
            continue
        r.update(map_activity=act, map_split=split, conf=old.get("confidence") or r["conf"])
        if not act:
            r["rev_status"] = "UNASSIGNED"
            r["rev_note"] = "Still no activity from the previous round."
        elif kind == "changed":
            r["rev_status"] = "CHANGED"
            r["rev_note"] = (f"Amount {old['amt']:,.2f} -> {r['amount']:,.2f} ({r['amount'] - old['amt']:+,.2f}). "
                             "Decision kept - check it still holds.")
        else:
            r["rev_status"] = "UNCHANGED"
    return deleted

def write_revised_mapping(results, path):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(MAP_COLS + REV_COLS)
        for r in results:
            w.writerow([r["ref"], r["row"], r["title"], r["desc"][:200], round(r["amount"], 2), r["map_activity"],
                        r["map_split"], r["conf"], r["basis"], " | ".join(r["flags"]), r["rev_status"], r["prev_amt"],
                        r["prev_act"], r["prev_split"], r["rev_note"]])

def write_revision(results, deleted, prev, path):
    """What changed between the two BOQs, biggest money first, with a reconciliation of the totals."""
    wb = Workbook(); wb.remove(wb.active)
    old_total, new_total = sum(x["amt"] for x in prev), sum(r["amount"] for r in results)
    n = lambda s: sum(1 for r in results if r["rev_status"] == s)
    move = lambda *st: sum(r["amount"] - (r["prev_amt"] or 0) for r in results if r["rev_status"] in st)
    gone = sum(x["amt"] for x in deleted)
    added, changed = move("NEW"), move("CHANGED", "REWORDED", "SCHEDULE_CHANGED", "UNASSIGNED", "UNCHANGED")
    check = old_total + added + changed - gone
    review = [r for r in results if r["rev_status"] in REVIEW_STATUS]
    ok = abs(check - new_total) < 0.01
    sheet(wb, "Summary", ["Item", "Value"], [
        ["Previous BOQ lines", len(prev)], ["Revised BOQ lines", len(results)],
        ["Decisions carried with no review", n("UNCHANGED")], ["Lines to review", len(review)],
        ["   New", n("NEW")], ["   Amount changed (decision kept)", n("CHANGED")],
        ["   Reworded or moved (matched fresh)", n("REWORDED")],
        ["   Schedule changed under the decision", n("SCHEDULE_CHANGED")],
        ["   Never assigned", n("UNASSIGNED")], ["Deleted lines", len(deleted)],
        ["Previous total", round(old_total, 2)], ["Added by new lines", round(added, 2)],
        ["Change on existing lines", round(changed, 2)], ["Removed by deleted lines", round(-gone, 2)],
        ["Revised total", round(new_total, 2)],
        ["Reconciliation", "PASS" if ok else f"FAIL - off by {new_total - check:,.2f}"],
    ], [42, 24])
    money = lambda r: -abs(r["amount"] - (r["prev_amt"] or 0))
    sheet(wb, "To Review", ["Status", "BOQ ref", "Heading", "Description", "Previous amount", "Amount", "Activity",
                            "Split", "Previous activity", "Note"],
          [[r["rev_status"], r["ref"], r["title"], r["desc"], r["prev_amt"], round(r["amount"], 2),
            r["map_activity"] or "UNMATCHED", r["map_split"], r["prev_act"], r["rev_note"]]
           for r in sorted(review, key=money)], [18, 8, 26, 50, 14, 14, 28, 16, 28, 60])
    sheet(wb, "Deleted", ["BOQ ref", "Heading", "Description", "Amount", "Was on activity"],
          [[x["ref"], x["heading"], x["description"], x["amt"], x["activity_group"]]
           for x in sorted(deleted, key=lambda x: -x["amt"])], [8, 26, 50, 14, 28])
    wb.save(path)
    return len(review), len(deleted), n("UNCHANGED"), ok

# ------------------------------------------------------------------ allocation
def allocate(results, groups, mapping=None):
    lines = []
    for r in results:
        m = (mapping or {}).get(str(r["row"]))
        if m:
            r["group"] = m["group"]
            targets = [m["single"]] if m["single"] else (groups.get(m["group"]) or [])
            split = m["split"]
        else:
            targets = groups.get(r["group"]) or []
            split = "equal"
        if not targets:
            r["alloc"] = []
            continue
        zk = zone_key(r.get("zone_hint"))
        on_floor = [a for a in targets if zk and zone_key(a["zone"]) == zk]
        if (split.lower() == "equal" or not split) and on_floor:
            # the BOQ says which floor this line is for ("4TH FLOOR") - load it there, not everywhere
            targets, split = on_floor, "equal"
            r["split_rule"] = f"BOQ floor heading '{r['zone_hint']}' -> {on_floor[0]['code']}"
            weights = {a["code"]: 1 for a in targets}
        elif split.lower() == "equal" or not split:
            weights = {a["code"]: 1 for a in targets}
            r["split_rule"] = (f"equal across {len(targets)} zones where this activity exists"
                               if len(targets) > 1 else "single activity")
        else:
            wz = {}
            for part in split.split(";"):
                if ":" in part:
                    z, v = part.split(":", 1)
                    wz[z.strip().lower()] = float(v)
            if any(not math.isfinite(v) or v < 0 for v in wz.values()):
                sys.exit(f"BOQ line {r['ref']}: split weights must be finite and non-negative")
            weights = {a["code"]: wz.get(a["zone"].lower(), wz.get(a["code"].lower(), 0)) for a in targets}
            r["split_rule"] = f"custom split {split}"
        tot = sum(weights.values())
        if tot <= 0:
            sys.exit(f"BOQ line {r['ref']}: split '{split}' matches none of {[a['zone'] for a in targets]}")
        r["alloc"] = []
        running = 0.0
        tlist = [a for a in targets if weights[a["code"]] > 0]
        for i, a in enumerate(tlist):
            share = weights[a["code"]] / tot
            cost = round(r["amount"] * share, 2) if i < len(tlist) - 1 else round(r["amount"] - running, 2)
            running += cost
            qty = round(r["qty"] * share, 4) if r["qty"] is not None else None
            r["alloc"].append({"act": a, "share": share, "cost": cost, "qty": qty})
        if len(tlist) > 1 and split.lower() == "equal":
            r["split_rule"] = (f"equal across {', '.join(a['zone'] or a['code'] for a in tlist)} "
                               "- edit the split in the mapping file if zones differ")
    return results

# ------------------------------------------------------------------ outputs
EXP_FIELDS = ["cost_item_id", "acct_id", "pobs_id", "cost_type_id", "proj_id", "task_id", "cost_name",
              "po_number", "vendor_name", "act_cost", "cost_per_qty", "remain_cost", "target_cost",
              "cost_load_type", "auto_compute_act_flag", "target_qty", "qty_name", "cost_descr",
              "contract_manager_import"]

def expense_name(r):
    base = r["title"] or r["desc"]
    return norm(f"BOQ {r['ref']} {base}")[:100]

def expense_rows(results, mode="line"):
    """One P6 expense per BOQ line per activity ('line'), or - for big BOQs - one per BOQ heading per
    activity ('heading'), so a 1,300-line BOQ gives a reviewable few hundred expenses, not 19,000.
    Grouped rows keep a quantity only when every line in the group has the same unit."""
    out = OrderedDict()
    for r in results:
        for al in r.get("alloc", []):
            if mode == "line":
                k = (al["act"]["code"], r["ref"])
                name, desc = expense_name(r), norm(r["desc"])
            else:
                head = r["title"] or r["desc"]
                k = (al["act"]["code"], (r.get("division") or "").lower(), head.lower())
                dv = div_number(r.get("division"))
                name, desc = norm(f"BOQ {'Div ' + str(dv) + ' ' if dv else ''}{head}")[:100], ""
            e = out.setdefault(k, {"act": al["act"], "name": name, "cost": 0.0, "qty": 0.0,
                                   "units": set(), "desc": desc, "lines": 0})
            e["cost"] += al["cost"]; e["lines"] += 1
            e["units"].add((r["unit"] or "").strip().lower())
            e["qty"] += al["qty"] or 0
    rows = []
    for e in out.values():
        same_unit = len(e["units"]) == 1 and e["qty"] > 0
        qty = e["qty"] if same_unit else 1
        unit = (next(iter(e["units"])) if same_unit else "Lot") or "Each"
        desc = e["desc"] or f"{e['lines']} BOQ line(s)"
        rows.append({"act": e["act"], "name": e["name"], "cost": round(e["cost"], 2), "qty": qty,
                     "unit": unit[:40], "desc": desc[:250]})
    return rows

def write_xer(xer, results, out_path, mode="line"):
    """Original XER, untouched, plus a PROJCOST table inserted before %E."""
    rows = []
    cid = 900000
    for e in expense_rows(results, mode):
        cid += 1
        per = e["cost"] / e["qty"] if e["qty"] else e["cost"]
        rows.append([str(cid), "", "", "", xer.project["proj_id"], e["act"]["task_id"],
                     e["name"], "", "", "0.0000", f"{per:.4f}", f"{e['cost']:.4f}",
                     f"{e['cost']:.4f}", "CL_Uniform", "N", f"{e['qty']:.4f}", e["unit"], e["desc"], ""])
    if "PROJCOST" in xer.tables:
        sys.exit("This XER already has expenses (PROJCOST). Remove them or export without expenses.")
    body = ["%T\tPROJCOST", "%F\t" + "\t".join(EXP_FIELDS)]
    for r in rows:
        body.append("%R\t" + "\t".join(v.replace("\t", " ") for v in r))
    end = re.search(rb"(?m)^%E[ \t]*(?:\r?\n)?\Z", xer.raw)
    if not end:
        sys.exit(f"{xer.path}: missing final %E record")
    newline = b"\r\n" if b"\r\n" in xer.raw else b"\n"
    inserted = ("\n".join(body) + "\n").encode(xer.encoding, errors="replace").replace(b"\n", newline)
    if end.start() and xer.raw[end.start() - 1:end.start()] not in (b"\n", b"\r"):
        inserted = newline + inserted
    Path(out_path).write_bytes(xer.raw[:end.start()] + inserted + xer.raw[end.start():])
    return len(rows)

def write_p6_xlsx(results, out_path, mode="line"):
    """P6 spreadsheet-import layout: row 1 = field names, row 2 = column labels (as P6 exports)."""
    wb = Workbook()
    ws = wb.active; ws.title = "PROJCOST"
    ws.append(["task_id", "cost_name", "target_cost", "target_qty", "qty_name", "cost_per_qty",
               "delete_record_flag"])
    ws.append(["Activity ID", "Expense Item", "Budgeted Cost", "Budgeted Units", "Unit of Measure",
               "Price / Unit", "Delete This Row"])
    for e in expense_rows(results, mode):
        ws.append([e["act"]["code"], e["name"], e["cost"], round(e["qty"], 4), e["unit"],
                   round(e["cost"] / e["qty"], 4) if e["qty"] else e["cost"], None])
    ud = wb.create_sheet("USERDATA")
    ud.append(["user_data"]); ud.append(["UserSettings Do Not Edit"])
    ud.append(["DurationQtyType=QT_Hour\nShowAsPercentage=0\nSmallScaleQtyType=QT_Hour\n"
               "DateFormat=dd/mm/yyyy\nCurrencyFormat=US Dollar\n"])
    wb.save(out_path)

HEAD = Font(bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor="1F3A5F")
FILL = {"HIGH": "D9EAD3", "MEDIUM": "FFF2CC", "LOW": "F4CCCC"}

def sheet(wb, title, header, rows, widths):
    ws = wb.create_sheet(title)
    ws.append(header)
    for c in ws[1]:
        c.font, c.fill = HEAD, HFILL
    for r in rows:
        ws.append(r)
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    return ws

def write_review(results, groups, boq_total, xer, boq_path, out_path, answer=None):
    wb = Workbook(); wb.remove(wb.active)
    loaded = sum(al["cost"] for r in results for al in r.get("alloc", []))
    unmatched = [r for r in results if not r.get("alloc")]
    used = {al["act"]["code"] for r in results for al in r.get("alloc", [])}
    empty = [a for g in groups.values() for a in g if a["code"] not in used]
    flagged = [r for r in results if r["flags"]]
    counts = {c: sum(1 for r in results if r["conf"] == c) for c in ("HIGH", "MEDIUM", "LOW")}
    summ = [
        ["BOQ file", str(boq_path)], ["Schedule", f"{xer.project['proj_short_name']} ({xer.path.name})"],
        ["BOQ lines", len(results)], ["BOQ total", round(boq_total, 2)],
        ["Loaded into P6", round(loaded, 2)],
        ["Difference", round(boq_total - loaded, 2)],
        ["Totals check", "PASS - every BOQ cost is loaded once" if not unmatched and
         all(abs(r["amount"] - sum(a["cost"] for a in r["alloc"])) < 0.005 for r in results)
         else "FAIL - see Unmatched lines"],
        ["Matches", f"{counts['HIGH']} HIGH / {counts['MEDIUM']} MEDIUM / {counts['LOW']} LOW"],
        ["Lines with flags", len(flagged)], ["Unmatched BOQ lines", len(unmatched)],
        ["Activities with no BOQ cost", len(empty)],
    ]
    if answer:
        summ += [["Answer-key check", answer["verdict"]]]
    ws = sheet(wb, "Summary", ["Item", "Value"], summ, [30, 90])
    sheet(wb, "Matches",
          ["BOQ ref", "Row", "Heading", "Description", "Unit", "Qty", "Amount", "Activity", "Matched on",
           "Score", "Confidence", "Split", "Flags"],
          [[r["ref"], r["row"], r["title"], r["desc"], r["unit"], r["qty"], round(r["amount"], 2),
            (groups[r["group"]][0]["base"] if r.get("group") in groups else "UNMATCHED"), r["basis"],
            r["score"], r["conf"], r.get("split_rule", ""), "\n".join(r["flags"])] for r in results],
          [8, 6, 24, 50, 7, 9, 13, 24, 14, 7, 11, 28, 60])
    m = wb["Matches"]
    for row in m.iter_rows(min_row=2):
        row[10].fill = PatternFill("solid", fgColor=FILL[row[10].value])
    by_act = defaultdict(float)
    for r in results:
        for al in r.get("alloc", []):
            by_act[al["act"]["code"]] += al["cost"]
    acts = [a for g in groups.values() for a in g]
    sheet(wb, "Activity Costs", ["Activity ID", "Activity name", "Zone", "WBS", "BOQ cost loaded"],
          [[a["code"], a["name"], a["zone"], " > ".join(a["wbs"]), round(by_act.get(a["code"], 0), 2)]
           for a in sorted(acts, key=lambda a: a["code"])], [14, 30, 16, 50, 16])
    hr = OrderedDict()
    for r in results:
        k = (r.get("division") or "", r["title"] or "(no heading)",
             groups[r["group"]][0]["base"] if r.get("group") in groups else "UNMATCHED - assign")
        e = hr.setdefault(k, {"n": 0, "amt": 0.0, "conf": set(), "flags": set(), "rule": set()})
        e["n"] += 1; e["amt"] += r["amount"]; e["conf"].add(r["conf"])
        e["flags"].update(re.sub(r"\d[\d,.]*", "#", f) for f in r["flags"])
        if r.get("split_rule"):
            e["rule"].add(r["split_rule"].split(" - ")[0])
    order = {"LOW": 0, "MEDIUM": 1, "HIGH": 2}
    rows_hr = sorted(hr.items(), key=lambda kv: (min(order[c] for c in kv[1]["conf"]), -kv[1]["amt"]))
    ws_hr = sheet(wb, "Heading Review",
                  ["Division", "BOQ heading", "Loaded onto activity", "Lines", "Amount", "Confidence",
                   "How split", "Flags (see Matches for each line)"],
                  [[d, h, a, e["n"], round(e["amt"], 2), "/".join(sorted(e["conf"], key=order.get)),
                    "; ".join(sorted(e["rule"]))[:120], "\n".join(sorted(e["flags"]))[:500]]
                   for (d, h, a), e in rows_hr], [22, 36, 36, 7, 15, 12, 40, 70])
    for row in ws_hr.iter_rows(min_row=2):
        row[5].fill = PatternFill("solid", fgColor=FILL[row[5].value.split("/")[0]])
    wb.move_sheet("Heading Review", offset=-(len(wb.sheetnames) - 2))
    sheet(wb, "Flags", ["BOQ ref", "Heading", "Flag"],
          [[r["ref"], r["title"], f] for r in results for f in r["flags"]]
          + [["-", a["name"], "Activity has no BOQ cost - missing BOQ line, or intentionally unpriced?"]
             for a in sorted(empty, key=lambda a: a["code"])], [8, 30, 100])
    if answer:
        sheet(wb, "Answer Key", ["Activity ID", "Planner (manual)", "This tool", "Difference", "Match"],
              answer["rows"], [14, 18, 18, 14, 8])
    sheet(wb, "Method", ["Step", "What happens"], [
        ["1 Read", "BOQ header row found automatically (description, heading, unit, qty, rate, amount). "
                   "Planner-added mapping columns (activity, floor, %) are ignored - the tool never peeks."],
        ["2 Match", "Each BOQ heading is matched to a schedule activity by trade words, tolerant of "
                    "misspellings and synonyms (config/synonyms.json). Description is checked too - "
                    "if it points somewhere else, the line is flagged."],
        ["3 Split", "If the activity repeats per floor/zone, the cost splits equally across the zones "
                    "where it exists (or by your custom split in the mapping file)."],
        ["4 Check", "Loaded total must equal the BOQ total to the cent. Unmatched lines and unpriced "
                    "activities are listed."],
        ["5 Write", "Expenses go into P6 as one expense item per BOQ line per activity, with quantity, "
                    "unit and unit price, so cost AND quantity are loaded."],
        ["Confidence", "HIGH = heading matches the activity exactly. MEDIUM = partial/fuzzy match or a "
                       "flag. LOW = no clear activity - assign by hand."],
    ], [14, 110])
    wb.save(out_path)
    return loaded, unmatched, empty, flagged

# ------------------------------------------------------------------ answer key (testing only)
def check_answer(results, key_path, sheet_name, id_col, cost_col):
    wb = load_workbook(key_path, data_only=True, read_only=True)
    ws = wb[sheet_name]
    truth = {}
    for r in ws.iter_rows(values_only=True):
        if r and isinstance(r[id_col], str) and isinstance(r[cost_col], (int, float)) \
                and not r[id_col].lower().startswith("grand"):
            truth[r[id_col]] = float(r[cost_col])
    mine = defaultdict(float)
    for x in results:
        for al in x.get("alloc", []):
            mine[al["act"]["code"]] += al["cost"]
    rows, ok = [], 0
    for code in sorted(set(truth) | set(mine)):
        t, m = truth.get(code, 0.0), mine.get(code, 0.0)
        good = abs(t - m) <= 0.02  # cents rounding only
        ok += good
        rows.append([code, round(t, 2), round(m, 2), round(m - t, 2), "YES" if good else "NO"])
    verdict = f"{ok}/{len(rows)} activities match the planner's manual cost loading"
    return {"rows": rows, "verdict": verdict, "ok": ok, "n": len(rows)}

# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser(description="Load BOQ cost + quantity into a Primavera P6 schedule.")
    ap.add_argument("mode", choices=["match", "build", "run", "revise"])
    ap.add_argument("boq"); ap.add_argument("xer")
    ap.add_argument("--sheet", help="BOQ sheet name (default: first sheet)")
    ap.add_argument("--out", default="outputs/cost-load")
    ap.add_argument("--previous", help="revise: the approved mapping CSV of the previous BOQ")
    ap.add_argument("--answer-key", help="TEST ONLY: workbook with a planner's manual result")
    ap.add_argument("--answer-sheet", default="Cost To P6")
    ap.add_argument("--expenses", choices=["auto", "line", "heading"], default="auto",
                    help="one P6 expense per BOQ line, or per BOQ heading per activity "
                         "(auto: per heading when the BOQ has more than 300 lines)")
    a = ap.parse_args()

    out = Path(a.out); out.parent.mkdir(parents=True, exist_ok=True)
    xer = Xer(a.xer)
    groups = group_activities(xer.activities())
    items, hdr = read_boq(a.boq, a.sheet)
    syn = load_synonyms()
    results = match_items(items, groups, syn, load_csi())
    rate_outliers(results); duplicate_lines(results)
    map_path = Path(str(out) + "-mapping.csv")

    if a.mode == "revise":
        if not a.previous:
            sys.exit("revise needs --previous <approved mapping.csv>.")
        if Path(a.previous).resolve() == map_path.resolve():
            sys.exit("--out would overwrite the previous mapping; use a new name, e.g. <job>-revB.")
        prev = read_previous(a.previous)
        deleted = revise(results, groups, prev)
        write_revised_mapping(results, map_path)
        rev_path = str(out) + "-revision.xlsx"
        review, gone, kept, ok = write_revision(results, deleted, prev, rev_path)
        print(f"{kept} decisions carried, {review} lines to review, {gone} deleted. "
              f"Reconciliation {'PASS' if ok else 'FAIL'}.")
        print(f"Review: {rev_path}   edit: {map_path}   then run: build")
        return
    if a.mode == "match":
        write_mapping(results, groups, map_path)
        c = {k: sum(1 for r in results if r["conf"] == k) for k in ("HIGH", "MEDIUM", "LOW")}
        print(f"{len(items)} BOQ lines, {sum(len(g) for g in groups.values())} activities "
              f"({len(groups)} distinct). Matches: {c['HIGH']} HIGH, {c['MEDIUM']} MEDIUM, {c['LOW']} LOW.")
        print(f"Review and edit: {map_path}   then run: build")
        return
    mapping = None
    if a.mode == "build":
        if not map_path.exists():
            sys.exit(f"No mapping file at {map_path}. Run 'match' first.")
        mapping = read_mapping(map_path, groups)
        expected = {str(r["row"]) for r in results}
        if set(mapping) != expected:
            sys.exit(f"{map_path}: mapping rows differ from BOQ; rerun 'match' for this workbook.")
        for r in results:
            m = mapping[str(r["row"])]
            # compare as numbers: Excel re-saves 14469.0 as 14469 (or 14,469) when a planner edits the CSV
            amt = _num(m["amount"])
            if str(m["ref"]).strip() != str(r["ref"]).strip() or amt is None or abs(amt - r["amount"]) > 0.01:
                sys.exit(f"{map_path}: BOQ line {r['ref']} changed since 'match'; rerun 'match'.")
    allocate(results, groups, mapping)
    boq_total = sum(r["amount"] for r in results)
    answer = check_answer(results, a.answer_key, a.answer_sheet, 0, 2) if a.answer_key else None
    loaded, unmatched, empty, flagged = write_review(results, groups, boq_total, xer, a.boq,
                                                      str(out) + "-review.xlsx", answer)
    mode = a.expenses if a.expenses != "auto" else ("heading" if len(results) > 300 else "line")
    n = write_xer(xer, results, str(out) + ".xer", mode)
    write_p6_xlsx(results, str(out) + "-p6-import.xlsx", mode)
    print(f"BOQ total {boq_total:,.2f} | loaded {loaded:,.2f} | diff {boq_total - loaded:,.2f}")
    print(f"{n} expense rows ({mode} level) -> {out}.xer and {out}-p6-import.xlsx")
    print(f"{len(flagged)} lines flagged, {len(unmatched)} unmatched, {len(empty)} activities with no cost")
    if answer:
        print("Answer key:", answer["verdict"])
    print(f"Review: {out}-review.xlsx")

if __name__ == "__main__":
    main()
