#!/usr/bin/env python3
"""
BOQ-reader check on a large real BOQ: planningengineer.net "How to read project BOQ from a Planning
Point of view" - a 17-floor tower, 22 division files (BOQ.zip, ~108.7M SAR).

The planner's own cleaned version ("Data Loading.xlsx", divisions 01-21, 444 lines) is the answer key:
every priced line + the "Main Title" the planner gave it. We read the RAW division files and compare.

    py tests/boq_reader_check.py [--show-misses]
Env: BIG_BOQ_DIR (default ../p6-test/hany), BIG_KEY (default ../Data Loading.xlsx)
"""
import os, re, sys
from collections import Counter, defaultdict
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import boq_to_p6 as E  # noqa: E402

BOQ_DIR = Path(os.environ.get("BIG_BOQ_DIR", ROOT.parent / "p6-test" / "hany"))
KEY = Path(os.environ.get("BIG_KEY", ROOT.parent / "Data Loading.xlsx"))

def divnum(s):
    m = re.search(r"(?:div(?:ision)?\.?\s*)(\d+)", str(s), re.I)
    return int(m.group(1)) if m else None

def same_title(a, b):
    ta, tb = set(E.tokens(a)), set(E.tokens(b))
    if not ta or not tb:
        return False
    return E.fuzzy_overlap(list(ta), list(tb)) >= 0.5 or E.fuzzy_overlap(list(tb), list(ta)) >= 0.5

def main():
    if not (BOQ_DIR.exists() and KEY.exists()):
        print(f"SKIP: need {BOQ_DIR} and {KEY}")
        return 0
    items, _ = E.read_boq(str(BOQ_DIR))
    mine = defaultdict(list)
    for it in items:
        d = divnum(it["division"]) or divnum(it["file"])
        mine[1 if d == 2 else d].append(it)  # file "02GENERAL REQUIREMENT" = Division 01

    wb = load_workbook(KEY, data_only=True, read_only=True)
    rows = list(wb["BOQ"].iter_rows(values_only=True))
    hdr = rows[0]
    key = defaultdict(list)
    for r in rows[1:]:
        d = dict(zip(hdr, r))
        if d.get("AMOUNT") in (None, ""):
            continue
        dn = divnum(d["Div No."])
        key[1 if dn == 2 else dn].append(d)  # the key labels Div 01 as "Div 2"

    total_key = total_mine = lines_ok = titles_ok = titles_n = 0
    misses = []
    print(f"{'Div':>4} {'key lines':>9} {'read':>5} {'key total':>14} {'read total':>14}  lines  titles")
    for dv in sorted(key):
        k, m = key[dv], mine.get(dv, [])
        kt = sum(float(x["AMOUNT"]) for x in k); mt = sum(x["amount"] for x in m)
        total_key += kt; total_mine += mt
        # pair lines by amount (+qty) - order-independent
        pool = defaultdict(list)
        for x in m:
            pool[(round(x["amount"], 2), round(x["qty"] or 0, 3))].append(x)
        matched = 0; tok = 0
        for x in k:
            cand = pool.get((round(float(x["AMOUNT"]), 2), round(float(x["QTY."] or 0), 3)))
            if cand:
                got = cand.pop(0); matched += 1
                if same_title(x["Main Title"], got["title"]):
                    tok += 1
                else:
                    misses.append((dv, str(x["Main Title"]).strip(), got["title"], got["desc"][:50]))
        lines_ok += matched; titles_ok += tok; titles_n += matched
        print(f"{dv:>4} {len(k):>9} {len(m):>5} {kt:>14,.0f} {mt:>14,.0f}  {matched:>3}/{len(k):<3} "
              f"{tok:>3}/{matched:<3}")
    n = sum(len(v) for v in key.values())
    print(f"\nLines found:   {lines_ok}/{n}")
    print(f"Totals:        key {total_key:,.2f} | read {total_mine:,.2f} | diff {total_mine - total_key:,.2f}")
    print(f"Headings:      {titles_ok}/{titles_n} agree with the planner's Main Title")
    other = {d: len(v) for d, v in mine.items() if d not in key}
    print(f"Beyond the key (Div 22-31, no answer key): {sum(other.values())} lines, "
          f"{sum(x['amount'] for d in other for x in mine[d]):,.0f}")
    print(f"Whole BOQ read: {len(items)} lines, {sum(x['amount'] for x in items):,.2f}")
    if "--show-misses" in sys.argv:
        print("\nHeading disagreements (div | planner | ours | line):")
        for m in misses:
            print("  ", m)
    return 0

if __name__ == "__main__":
    sys.exit(main())
