#!/usr/bin/env python3
"""
Regression check: run boq_to_p6.py on the case study and assert known-correct answers.

    py tests/regression.py

Test data = planningengineer.net case study "How to link the BOQ and P6 submitted baseline schedule"
(Hassan Reda, Jan 2025): NEWPROJ-4.xer + Case Study BOQ.xlsx, looked up in TEST_DIR
(default: ../p6-test/hassan next to this repo). The BOQ workbook also holds the planner's manual
result ("Cost To P6" sheet) - the answer key. Missing data = SKIP, not fail.
Expected values below were checked by hand against the BOQ and against P6 after import
(BOQ 1: 168 m3 x 86.125 = 14,469.00 shown on the imported expense in P6 Professional 25).
"""
import os, subprocess, sys, tempfile
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
TEST_DIR = Path(os.environ.get("TEST_DIR", ROOT.parent / "p6-test" / "hassan"))
ENGINE = ROOT / "scripts" / "boq_to_p6.py"
BOQ, XER = TEST_DIR / "Case Study BOQ.xlsx", TEST_DIR / "NEWPROJ-4.xer"

def rows(wb, tab):
    data = list(wb[tab].iter_rows(values_only=True))
    return [dict(zip(data[0], r)) for r in data[1:]]

def xer_expenses(path):
    fields, out, cur = None, [], None
    for ln in open(path, encoding="cp1252"):
        p = ln.rstrip("\r\n").split("\t")
        if p[0] == "%T":
            cur = p[1]
        elif p[0] == "%F" and cur == "PROJCOST":
            fields = p[1:]
        elif p[0] == "%R" and cur == "PROJCOST":
            out.append(dict(zip(fields, p[1:])))
    return out

def checks(out):
    rev = load_workbook(str(out) + "-review.xlsx")
    summ = {r["Item"]: r["Value"] for r in rows(rev, "Summary")}
    yield "totals check PASS", str(summ["Totals check"]).startswith("PASS")
    yield "BOQ total 840,414", abs(summ["BOQ total"] - 840414) < 0.01
    yield "answer key 48/48", str(summ.get("Answer-key check", "")).startswith("48/48")
    m = {str(r["BOQ ref"]): r for r in rows(rev, "Matches")}
    yield "22 BOQ lines read (incl. 'TOTAL STEEL & WOODEN DOORS' item)", len(m) == 22
    yield "no LOW / unmatched lines", all(r["Confidence"] != "LOW" for r in m.values())
    f = lambda ref: str(m[ref]["Flags"] or "")
    yield "line 7: paint description under Doors flagged", "Painting" in f("7")
    yield "line 8: doors description under Windows flagged", "Doors" in f("8")
    yield "line 9: kitchen cabinets under WaterProofing flagged", "does not mention" in f("9")
    yield "line 10: handrails under wall tiles flagged", "does not mention" in f("10")
    yield "lines 11/12: 10x skirting rate flagged", "10x" in f("11") and "10x" in f("12")
    yield "lines 13/22: duplicate flagged", "duplicate" in f("13") and "duplicate" in f("22")
    yield "line 12 NOT flagged as sanitary (toilet = room)", "Sanitary" not in f("12")
    yield "line 7 NOT blamed on waterproofing (weatherproof)", "WaterProofing" not in f("7")
    clean = [r for k, r in m.items() if k not in {"7", "8", "9", "10", "11", "12", "13", "22"}]
    yield "the other 14 lines have no flags", all(not r["Flags"] for r in clean)
    exp = xer_expenses(str(out) + ".xer")
    yield "54 expense rows in XER", len(exp) == 54
    yield "XER expense total 840,414", abs(sum(float(e["target_cost"]) for e in exp) - 840414) < 0.01
    b1 = [e for e in exp if e["cost_name"] == "BOQ 1 Excavation"]
    yield "BOQ 1: 168 m3 @ 86.125 = 14,469", len(b1) == 1 and float(b1[0]["target_qty"]) == 168 \
        and float(b1[0]["target_cost"]) == 14469 and b1[0]["qty_name"] == "m3"
    raw = open(str(out) + ".xer", "rb").read()
    orig = open(XER, "rb").read()
    yield "original XER kept byte-identical", raw[:raw.find(b"%T\tPROJCOST")] == orig[:orig.rfind(b"%E")]
    yield "no blank/doubled lines in XER", b"\r\n\r\n" not in raw and b"\r\r" not in raw
    p6 = load_workbook(str(out) + "-p6-import.xlsx")
    yield "spreadsheet import has 54 rows + 2 header rows", p6["PROJCOST"].max_row == 56

TOWER_BOQ = Path(os.environ.get("BIG_BOQ_DIR", ROOT.parent / "p6-test" / "hany"))
TOWER_XER = ROOT.parent / "p6-test" / "tower" / "TOWER-17.xer"
FIXTURE = ROOT / "tests" / "fixtures" / "make_tower_schedule.py"

def tower_checks(tmp):
    """Real 22-file tower BOQ (1,332 lines) onto a 688-activity schedule. Expected values checked by
    hand in the review workbook and after importing the result into P6 Professional 25."""
    import csv, time
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    if not TOWER_XER.exists():
        subprocess.run([sys.executable, str(FIXTURE)], capture_output=True, env=env)
    out = Path(tmp) / "tower"
    t0 = time.time()
    r = subprocess.run([sys.executable, str(ENGINE), "run", str(TOWER_BOQ), str(TOWER_XER), "--out", str(out)],
                       capture_output=True, text=True, encoding="utf-8", env=env)
    secs = time.time() - t0
    if r.returncode != 0:
        yield f"tower run crashed: {r.stderr[-300:]}", False
        return
    rev = load_workbook(str(out) + "-review.xlsx")
    summ = {x["Item"]: x["Value"] for x in rows(rev, "Summary")}
    m = rows(rev, "Matches")
    yield f"tower: runs in under 2 minutes ({secs:.0f}s)", secs < 120
    yield "tower: 1,332 BOQ lines read", summ["BOQ lines"] == 1332
    yield "tower: BOQ total 106,996,378.52", abs(summ["BOQ total"] - 106996378.52) < 0.01
    yield "tower: at least 90% loaded", summ["Loaded into P6"] >= 0.90 * 106996378.52
    yield "tower: totals check honestly FAILs while lines are unassigned", str(summ["Totals check"]).startswith("FAIL")
    yield "tower: no LOW line carries money", all(x["Activity"] == "UNMATCHED" for x in m if x["Confidence"] == "LOW")
    act = {}
    for x in m:
        act.setdefault(str(x["Heading"]).strip(), set()).add(x["Activity"])
    expect = {"BUSWAY": "Main Cable Risers & Busbar", "FLOOR FINISHING": "Floor Finishes",
              "CEILING FINISHES": "False Ceilings", "BASE FINISHES": "Skirting & Transition Strips",
              "CABINETS": "Joinery & Casework", "Metal Doors": "Internal Doors & Ironmongery",
              "PIPES": "Sprinkler Pipework", "TRANSFORMERS": "RMU, MV Cables & Transformers",
              "NOVEC 1230 SYSTEM": "Clean Agent Fire Suppression (Server & Electrical Rooms)",
              "GUARD RAIL": "Stair Balustrades, Handrails & Metal Works"}
    for h, a in expect.items():
        yield f"tower: '{h}' -> '{a}'", act.get(h) == {a}
    raft = [x for x in m if x["Activity"] == "Raft Foundation Concrete"]
    yield "tower: only the 1500 mm raft line goes to the raft", len(raft) == 1 and raft[0]["Amount"] == 1900570
    duct = [x for x in m if str(x["Heading"]).strip() == "GALVANIZED STEEL DUCTS"]
    yield "tower: ambiguous ducts stay unloaded, right weak guess shown", duct and all(
        x["Activity"] == "UNMATCHED" and "HVAC Ductwork" in str(x["Flags"]) for x in duct)
    exp = xer_expenses(str(out) + ".xer")
    yield "tower: grouped expenses (under 2,000 rows, not 19,000)", 0 < len(exp) < 2000
    yield "tower: XER expense total = loaded total",  \
        abs(sum(float(e["target_cost"]) for e in exp) - summ["Loaded into P6"]) < 0.05
    raw, orig = open(str(out) + ".xer", "rb").read(), open(TOWER_XER, "rb").read()
    yield "tower: original schedule kept byte-identical", raw[:raw.find(b"%T\tPROJCOST")] == orig[:orig.rfind(b"%E")]
    # floor routing: this BOQ's per-floor NOVEC lines belong to one building-wide activity, so force
    # them onto a per-floor activity via the mapping file and check '4TH FLOOR' lands on floor 4 only
    fo = Path(tmp) / "floors"
    subprocess.run([sys.executable, str(ENGINE), "match", str(TOWER_BOQ), str(TOWER_XER), "--out", str(fo)],
                   capture_output=True, env=env)
    p = f"{fo}-mapping.csv"
    mp = list(csv.DictReader(open(p, encoding="utf-8-sig")))
    for x in mp:
        if x["heading"].strip() == "NOVEC 1230 SYSTEM":
            x["activity_group"] = "Fire Alarm Detectors & Devices"
    with open(p, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=mp[0].keys()); w.writeheader(); w.writerows(mp)
    subprocess.run([sys.executable, str(ENGINE), "build", str(TOWER_BOQ), str(TOWER_XER), "--out", str(fo)],
                   capture_output=True, env=env)
    fm = rows(load_workbook(f"{fo}-review.xlsx"), "Matches")
    f4 = [x for x in fm if str(x["Heading"]).strip() == "NOVEC 1230 SYSTEM" and "'4TH FLOOR'" in str(x["Split"])]
    yield "tower: '4TH FLOOR' lines routed to floor 4 only (F04)", f4 and all("-> F04" in x["Split"] for x in f4)
    costs = {x["Activity ID"]: x["BOQ cost loaded"] for x in rows(load_workbook(f"{fo}-review.xlsx"), "Activity Costs")}
    base = {x["Activity ID"]: x["BOQ cost loaded"] for x in rows(rev, "Activity Costs")}
    # F04-25 also carries its equal share of the building-wide fire-alarm devices - check what was ADDED
    yield "tower: floor 4 gains exactly the 4th-floor NOVEC money (239,105 + 420,010)", \
        abs(costs.get("F04-25", 0) - base.get("F04-25", 0) - 659115) < 0.01

def revision_checks():
    """Rev B of the case study: one line inserted at the top (every row shifts), one amount changed, one
    line reworded, one deleted, on top of a planner's own override in Rev A."""
    import csv, importlib.util
    spec = importlib.util.spec_from_file_location("engine", ENGINE)
    eng = importlib.util.module_from_spec(spec); spec.loader.exec_module(eng)
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    with tempfile.TemporaryDirectory() as tmp:
        a, b = str(Path(tmp) / "revA"), str(Path(tmp) / "revB")
        run = lambda *x: subprocess.run([sys.executable, str(ENGINE), *map(str, x)], capture_output=True, env=env,
                                        text=True, encoding="utf-8")
        run("match", BOQ, XER, "--out", a)
        p = a + "-mapping.csv"
        mp = list(csv.DictReader(open(p, encoding="utf-8-sig")))
        by = {x["ref"]: x for x in mp}
        groups = eng.group_activities(eng.Xer(XER).activities())
        multi = next(g for g in groups.values() if len(g) > 1)
        by["5"].update(activity_group=multi[0]["base"], split=f"{multi[0]['zone']}:100")  # planner's override
        with open(p, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=mp[0].keys()); w.writeheader(); w.writerows(mp)
        wb = load_workbook(BOQ, data_only=True)
        ws = next(s for s in wb.worksheets if eng.find_header([list(r) for r in s.values])[0] is not None)
        data = [list(r) for r in ws.values]
        hdr, cols = eng.find_header(data)
        ref_row = lambda ref: next(i + 1 for i, r in enumerate(data) if i > hdr and
                                   (str(r[cols["ref"]]).strip() if "ref" in cols else str(i + 1)) == ref)
        cell = lambda row, k: ws.cell(row, cols[k] + 1)
        old2 = float(by["2"]["amount"])
        cell(ref_row("2"), "qty").value *= 2
        if "amount" in cols:
            cell(ref_row("2"), "amount").value = old2 * 2
        cell(ref_row("3"), "desc").value = str(cell(ref_row("3"), "desc").value) + " incl."
        ws.delete_rows(ref_row("4"))
        ws.insert_rows(hdr + 2)
        for k, v in dict(ref="N1", title="Stone Cladding", desc="Granite stone cladding", unit="m2", qty=10,
                         rate=100, amount=1000).items():
            if k in cols:
                cell(hdr + 2, k).value = v
        rev = Path(tmp) / "revB.xlsx"; wb.save(rev)
        r = run("revise", rev, XER, "--previous", p, "--out", b)
        yield f"revise runs ({r.stderr[-200:]})", r.returncode == 0
        if r.returncode:
            return
        new = {x["ref"]: x for x in csv.DictReader(open(b + "-mapping.csv", encoding="utf-8-sig"))}
        st = lambda ref: new[ref]["revision_status"]
        yield "22 + 1 - 1 lines", len(new) == 22
        yield "inserted line = NEW", st("N1") == "NEW"
        yield "doubled line = CHANGED, decision kept", st("2") == "CHANGED" and \
            new["2"]["activity_group"] == by["2"]["activity_group"]
        yield "reworded line = REWORDED, previous decision shown", st("3") == "REWORDED" and \
            new["3"]["previous_activity"] == by["3"]["activity_group"]
        yield "planner override survives shifted rows", st("5") == "UNCHANGED" and \
            new["5"]["activity_group"] == multi[0]["base"] and new["5"]["split"] == by["5"]["split"]
        same = [x for k, x in new.items() if k not in {"N1", "2", "3"}]
        yield "every other line UNCHANGED", all(x["revision_status"] == "UNCHANGED" for x in same)
        rv = load_workbook(b + "-revision.xlsx")
        summ = {x["Item"]: x["Value"] for x in rows(rv, "Summary")}
        yield "totals reconcile", summ["Reconciliation"] == "PASS"
        yield "3 lines to review", len(rows(rv, "To Review")) == 3
        yield "deleted line listed", [str(x["BOQ ref"]) for x in rows(rv, "Deleted")] == ["4"]
        r = run("build", rev, XER, "--out", b)
        yield f"build accepts revised mapping ({r.stderr[-200:]})", r.returncode == 0
        exp = xer_expenses(b + ".xer") if r.returncode == 0 else []
        yield "override loads on one activity in the XER", \
            len([e for e in exp if e["cost_name"].startswith("BOQ 5 ")]) == 1
        # a schedule that lost the override's activity: the decision must not load silently
        mp2 = list(csv.DictReader(open(p, encoding="utf-8-sig")))
        for x in mp2:
            if x["ref"] == "5":
                x["activity_group"] = "Activity Removed From Schedule"
        p2 = str(Path(tmp) / "old-mapping.csv")
        with open(p2, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=mp2[0].keys()); w.writeheader(); w.writerows(mp2)
        c = str(Path(tmp) / "revC")
        run("revise", rev, XER, "--previous", p2, "--out", c)
        n5 = next(x for x in csv.DictReader(open(c + "-mapping.csv", encoding="utf-8-sig")) if x["ref"] == "5")
        yield "removed activity = SCHEDULE_CHANGED, not carried", n5["revision_status"] == "SCHEDULE_CHANGED" and \
            n5["activity_group"] != "Activity Removed From Schedule"
        r = run("revise", rev, XER, "--previous", p, "--out", a)
        yield "refuses to overwrite the previous mapping", r.returncode != 0

def main():
    if not (BOQ.exists() and XER.exists()):
        print(f"SKIP: test data not found in {TEST_DIR}")
        sys.exit(0)
    passed = failed = 0
    if TOWER_BOQ.exists() and "--quick" not in sys.argv:
        with tempfile.TemporaryDirectory() as tmp:
            for label, good in tower_checks(tmp):
                if not good:
                    print(f"FAIL {label}")
                passed += bool(good); failed += (not good)
            print("  tower: done")
    else:
        print("SKIP tower: big BOQ not found (or --quick)")
    with tempfile.TemporaryDirectory() as tmp:
        for mode in ("run", "match+build"):
            out = Path(tmp) / mode.replace("+", "_")
            env = dict(os.environ, PYTHONIOENCODING="utf-8")
            steps = [["run"]] if mode == "run" else [["match"], ["build"]]
            ok = True
            for s in steps:
                cmd = [sys.executable, str(ENGINE), *s, str(BOQ), str(XER), "--out", str(out)]
                if s[0] != "match":
                    cmd += ["--answer-key", str(BOQ)]
                r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", env=env)
                if r.returncode != 0:
                    print(f"FAIL {mode}: engine crashed\n{r.stderr[-800:]}")
                    failed += 1; ok = False
                    break
            if not ok:
                continue
            for label, good in checks(out):
                if not good:
                    print(f"FAIL {mode}: {label}")
                passed += good; failed += (not good)
            print(f"  {mode}: done")
    # the mapping CSV is edited by planners in Excel, which re-saves 14469.0 as 14469 - build must accept it,
    # but must still refuse a mapping whose amounts really changed since 'match'
    import csv
    with tempfile.TemporaryDirectory() as tmp:
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        o = str(Path(tmp) / "xl")
        subprocess.run([sys.executable, str(ENGINE), "match", str(BOQ), str(XER), "--out", o], capture_output=True, env=env)
        def resave(fn):
            p = o + "-mapping.csv"
            mp = list(csv.DictReader(open(p, encoding="utf-8-sig")))
            fn(mp)
            with open(p, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.DictWriter(fh, fieldnames=mp[0].keys()); w.writeheader(); w.writerows(mp)
        build = lambda: subprocess.run([sys.executable, str(ENGINE), "build", str(BOQ), str(XER), "--out", o],
                                       capture_output=True, env=env).returncode
        resave(lambda mp: [x.update(amount=f"{float(x['amount']):,.0f}" if float(x["amount"]).is_integer()
                                    else x["amount"]) for x in mp])
        ok = build() == 0
        print("" if ok else "FAIL mapping re-saved by Excel (14,469) is accepted"); passed += ok; failed += not ok
        resave(lambda mp: mp[0].update(amount="15000"))
        ok = build() != 0
        print("" if ok else "FAIL mapping with a changed amount is refused"); passed += ok; failed += not ok
    for label, good in revision_checks():
        if not good:
            print(f"FAIL revise: {label}")
        passed += bool(good); failed += (not good)
    print(f"\n{passed} passed, {failed} failed")
    sys.exit(1 if failed else 0)

if __name__ == "__main__":
    main()
