#!/usr/bin/env python3
"""
TEST FIXTURE - builds a realistic P6 schedule (XER) for the 17-floor tower in the big-BOQ test,
because no real schedule was published for that BOQ.

Activity names follow normal planning practice (trade x floor, MEP first fix / second fix, plant
rooms as single activities - see MEP Guide for Planning & Scheduling, planningengineer.net).
They are written WITHOUT looking up BOQ headings, so the matcher is tested on vocabulary it was
not given. Uses the case-study XER (proven to import into P6 25) as the format template.

    py tests/fixtures/make_tower_schedule.py  -> ../p6-test/tower/TOWER-17.xer
"""
import datetime as dt
from collections import OrderedDict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
TEMPLATE = ROOT.parent / "p6-test" / "hassan" / "NEWPROJ-4.xer"
OUT = ROOT.parent / "p6-test" / "tower" / "TOWER-17.xer"

ZONES = ([("B3", "Lower Level 3"), ("B2", "Lower Level 2"), ("B1", "Lower Level 1"),
          ("GF", "Ground Floor"), ("MZ", "Mezzanine Floor")]
         + [(f"F{n:02d}", f"Level {n:02d}") for n in range(1, 18)] + [("RF", "Roof Level")])

# per-floor work, in site order: (name, days, discipline)
STRUCTURE = [("Columns & Core Walls Concrete", 8), ("Suspended Slab Concrete", 10)]
FLOOR_WORK = [
    ("Blockwork Partitions", 12, "Architectural"),
    ("Sprinkler Pipework", 10, "Mechanical"),
    ("Drainage & Water Supply Pipework", 10, "Mechanical"),
    ("HVAC Ductwork & Insulation", 14, "Mechanical"),
    ("Chilled Water Pipework", 10, "Mechanical"),
    ("Electrical Conduits & Cable Trays", 10, "Electrical"),
    ("Low Current Containment", 6, "Electrical"),
    ("Gypsum Drywall Partitions", 12, "Architectural"),
    ("Power Cabling & Wiring", 10, "Electrical"),
    ("Fire Alarm Cabling", 6, "Electrical"),
    ("Data & Telephone Cabling", 6, "Electrical"),
    ("False Ceilings", 10, "Architectural"),
    ("Floor Finishes", 12, "Architectural"),
    ("Wall Finishes & Painting", 12, "Architectural"),
    ("Skirting & Transition Strips", 4, "Architectural"),
    ("Internal Doors & Ironmongery", 6, "Architectural"),
    ("Joinery & Casework", 8, "Architectural"),
    ("Sanitary Fixtures & Fittings", 5, "Mechanical"),
    ("Fan Coil Units & VAV Boxes", 6, "Mechanical"),
    ("Diffusers, Grilles & Dampers", 5, "Mechanical"),
    ("Sprinkler Heads & Hose Cabinets", 4, "Mechanical"),
    ("Light Fittings", 5, "Electrical"),
    ("Switches & Socket Outlets", 4, "Electrical"),
    ("Distribution Boards", 3, "Electrical"),
    ("Fire Alarm Detectors & Devices", 4, "Electrical"),
    ("CCTV & Access Control Devices", 4, "Electrical"),
    ("Toilet Accessories & Cubicles", 3, "Architectural"),
]
BASEMENT_ONLY = [("Basement Retaining Walls", 10)]
# building-wide activities: (code, name, days, wbs, predecessor codes)
SINGLE = [
    ("G100", "Project Management & Submittals", 700, "General", []),
    ("G110", "Temporary Facilities & Site Services", 700, "General", []),
    ("G120", "Health, Safety & Site Security", 700, "General", []),
    ("G130", "Quality Control & Inspections", 700, "General", []),
    ("G140", "Site Cleaning & Waste Removal", 700, "General", []),
    ("E100", "Excavation, Shoring & Dewatering", 60, "Substructure", []),
    ("E110", "Blinding & Protection Screed", 10, "Substructure", ["E100"]),
    ("E120", "Raft Foundation Concrete", 30, "Substructure", ["E110"]),
    ("E130", "Below-Grade Damp-Proofing & Waterproofing", 25, "Substructure", ["S-B1-2"]),
    ("E140", "Backfilling & Compaction", 15, "Substructure", ["E130"]),
    ("A100", "Curtain Wall & Facade Glazing", 180, "Envelope", ["S-F10-2"]),
    ("A110", "Metal Cladding, Louvers & Screens", 90, "Envelope", ["S-F17-2"]),
    ("A120", "Precast Facade Panels", 40, "Envelope", ["S-F05-2"]),
    ("A130", "Entrance Canopies Steelwork & Metal Deck", 45, "Envelope", ["S-GF-2"]),
    ("A140", "Roof Waterproofing & Insulation", 25, "Envelope", ["S-RF-2"]),
    ("A150", "Stair Balustrades, Handrails & Metal Works", 60, "Architectural Common", ["S-RF-2"]),
    ("A160", "Wall Protection, Lockers & Specialties", 30, "Architectural Common", ["A150"]),
    ("A170", "Wayfinding & Traffic Signage", 20, "Architectural Common", ["A160"]),
    ("A180", "Vanity Tops & Furnishings", 15, "Architectural Common", ["A160"]),
    ("M100", "Chillers, Heat Exchangers & Primary Pumps", 60, "Central Plant", ["S-RF-2"]),
    ("M110", "Air Handling Units & Kitchen Exhaust", 60, "Central Plant", ["S-RF-2"]),
    ("M120", "Fire Pump Set & Jockey Pump", 20, "Central Plant", ["S-B3-2"]),
    ("M130", "Clean Agent Fire Suppression (Server & Electrical Rooms)", 45, "Central Plant", ["S-F17-2"]),
    ("M140", "Domestic Water Pumps, Tanks & Water Heaters", 30, "Central Plant", ["S-B3-2"]),
    ("M150", "Grey Water & Irrigation Systems", 30, "Central Plant", ["S-GF-2"]),
    ("M160", "Generator Fuel System", 20, "Central Plant", ["S-B1-2"]),
    ("M170", "Building Management System (BMS)", 90, "Central Plant", ["M100"]),
    ("L100", "RMU, MV Cables & Transformers", 45, "Main Electrical", ["S-B1-2"]),
    ("L110", "LV Switchgear & Main Distribution", 40, "Main Electrical", ["L100"]),
    ("L120", "Standby Generators & ATS", 30, "Main Electrical", ["L110"]),
    ("L130", "Main Cable Risers & Busbar", 60, "Main Electrical", ["L110"]),
    ("L140", "Earthing, Lightning Protection & Cathodic Protection", 30, "Main Electrical", ["S-RF-2"]),
    ("L150", "Main Telecom Room, Racks & Servers", 40, "Main Electrical", ["L130"]),
    ("L160", "Public Address, Nurse Call & AV Systems", 45, "Main Electrical", ["L150"]),
    ("L170", "Fire Alarm Main Panels & Graphics", 20, "Main Electrical", ["L150"]),
    ("L180", "Security Head-End (CCTV, Access, Intrusion)", 25, "Main Electrical", ["L150"]),
    ("V100", "Elevators & Lifts Installation", 150, "Vertical Transport", ["S-RF-2"]),
    ("V110", "Facade Access / Building Maintenance Unit", 40, "Vertical Transport", ["A140"]),
    ("X100", "Kitchen, Pantry & Nourishment Equipment", 20, "Fit-out Equipment", ["A160"]),
    ("X110", "Water Feature & Fountain", 20, "External Works", ["S-GF-2"]),
    ("X120", "External Paving, Landscaping & Wheel Stops", 45, "External Works", ["E140"]),
    ("C100", "Testing & Commissioning", 60, "Closeout", []),
    ("C110", "Closeout Documents & As-Builts", 30, "Closeout", ["C100"]),
]

def parse(path):
    lines = path.read_text(encoding="cp1252").splitlines()
    tables, fields, cur = OrderedDict(), {}, None
    for ln in lines[1:]:
        p = ln.split("\t")
        if p[0] == "%T":
            cur = p[1]; tables[cur] = []
        elif p[0] == "%F":
            fields[cur] = p[1:]
        elif p[0] == "%R":
            tables[cur].append(dict(zip(fields[cur], p[1:])))
    return lines[0], tables, fields

def main():
    header, tables, fields = parse(TEMPLATE)
    proj = dict(tables["PROJECT"][0])
    PID, CAL = "7717", proj["clndr_id"]
    proj.update(proj_id=PID, proj_short_name="TOWER-17", guid="", plan_start_date="2026-01-04 00:00",
                last_recalc_date="2026-01-04 00:00", scd_end_date="", last_schedule_date="",
                task_code_prefix="T")
    wbs_t = tables["PROJWBS"][0]; task_t = tables["TASK"][0]; pred_t = tables["TASKPRED"][0]

    wbs_rows, next_w = [], [800000]
    def wbs(name, short, parent, node=False):
        next_w[0] += 1
        r = dict(wbs_t); r.update(wbs_id=str(next_w[0]), proj_id=PID, proj_node_flag="Y" if node else "N",
                                   wbs_short_name=short, wbs_name=name, parent_wbs_id=parent, guid="",
                                   seq_num=str(len(wbs_rows)))
        wbs_rows.append(r); return r["wbs_id"]
    root = wbs("17-Floor Tower (BOQ load test)", "TOWER-17", wbs_t["parent_wbs_id"], node=True)
    W = {}
    for grp in ["General", "Substructure", "Superstructure", "Floors", "Envelope", "Architectural Common",
                "Central Plant", "Main Electrical", "Vertical Transport", "Fit-out Equipment",
                "External Works", "Closeout"]:
        W[grp] = wbs(grp, grp[:12].upper().replace(" ", ""), root)
    for code, zname in ZONES:
        W["S-" + code] = wbs(zname, code, W["Superstructure"])
        zf = wbs(zname, code, W["Floors"])
        for disc in ("Architectural", "Mechanical", "Electrical"):
            W[f"F-{code}-{disc}"] = wbs(disc, disc[:4].upper(), zf)

    tasks, preds, next_t = OrderedDict(), [], [900000]
    def task(code, name, days, wbs_id, after=(), mile=False):
        next_t[0] += 1
        tasks[code] = dict(task_t, task_id=str(next_t[0]), proj_id=PID, wbs_id=wbs_id, clndr_id=CAL,
                           task_code=code, task_name=name, guid="", status_code="TK_NotStart",
                           task_type="TT_FinMile" if mile else "TT_Task", days=days, after=list(after))
    task("M000", "Project Start", 0, W["General"], mile=True)
    prev = "E120"
    for code, _ in ZONES:
        a, b = f"S-{code}-1", f"S-{code}-2"
        if code.startswith("B"):
            task(f"S-{code}-0", f"{BASEMENT_ONLY[0][0]} - {code}", BASEMENT_ONLY[0][1], W["S-" + code], [prev])
            prev = f"S-{code}-0"
        task(a, f"{STRUCTURE[0][0]} - {code}", STRUCTURE[0][1], W["S-" + code], [prev])
        task(b, f"{STRUCTURE[1][0]} - {code}", STRUCTURE[1][1], W["S-" + code], [a])
        prev = b
        if code == "RF":
            continue
        chain = b
        for i, (nm, d, disc) in enumerate(FLOOR_WORK):
            c = f"{code}-{i + 1:02d}"
            task(c, f"{nm} - {code}", d, W[f"F-{code}-{disc}"], [chain])
            chain = c
    for code, nm, d, grp, after in SINGLE:
        task(code, nm, d, W[grp], after or ["M000"])
    last_floor = [f"{c}-{len(FLOOR_WORK):02d}" for c, _ in ZONES if c != "RF"]
    tasks["C100"]["after"] = last_floor + ["M170", "V100", "L180"]
    task("M999", "Building Handover", 0, W["Closeout"], ["C110"], mile=True)
    for c in ("E100",) + tuple(k for k, *_ in SINGLE[:5]):
        tasks[c]["after"] = ["M000"]

    # simple forward pass (7-day calendar, 8h days) so P6 opens it with sensible dates
    start = dt.datetime(2026, 1, 4, 8, 0); fin = {}
    def finish(code):
        if code in fin:
            return fin[code]
        t = tasks[code]
        s = max([finish(p) for p in t["after"]], default=start)
        f = s + dt.timedelta(days=t["days"])
        t["_s"], fin[code] = s, f
        return f
    for c in tasks:
        finish(c)
    fmt = lambda d: d.strftime("%Y-%m-%d %H:%M")
    task_rows = []
    for c, t in tasks.items():
        hrs = str(t["days"] * 8)
        s, f = fmt(t["_s"]), fmt(fin[c])
        r = {k: v for k, v in t.items() if k not in ("days", "after", "_s")}
        r.update(target_drtn_hr_cnt=hrs, remain_drtn_hr_cnt=hrs, total_float_hr_cnt="0", free_float_hr_cnt="0",
                 early_start_date=s, early_end_date=f, late_start_date=s, late_end_date=f,
                 target_start_date=s, target_end_date=f, restart_date=s, reend_date=f,
                 rem_late_start_date=s, rem_late_end_date=f, act_start_date="", act_end_date="",
                 cstr_date="", cstr_type="", create_date="2026-01-04 08:00", update_date="2026-01-04 08:00")
        task_rows.append(r)
        for p in t["after"]:
            q = dict(pred_t); q.update(task_pred_id=str(1900000 + len(preds)), task_id=t["task_id"],
                                       pred_task_id=tasks[p]["task_id"], proj_id=PID, pred_proj_id=PID,
                                       pred_type="PR_FS", lag_hr_cnt="0", aref="", arls="")
            preds.append(q)

    out = [header]
    for name, rows in tables.items():
        if name in ("UDFTYPE", "UDFVALUE"):
            continue
        rows = {"PROJECT": [proj], "PROJWBS": wbs_rows, "TASK": task_rows, "TASKPRED": preds}.get(name, rows)
        out.append("%T\t" + name)
        out.append("%F\t" + "\t".join(fields[name]))
        for r in rows:
            out.append("%R\t" + "\t".join(str(r.get(f, "")) for f in fields[name]))
    out.append("%E")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="cp1252", newline="") as fh:
        fh.write("\r\n".join(out) + "\r\n")
    n_work = sum(1 for t in task_rows if t["task_type"] == "TT_Task")
    print(f"{OUT}: {len(task_rows)} activities ({n_work} work + milestones), {len(wbs_rows)} WBS, "
          f"{len(preds)} links, finish {max(fin.values()):%Y-%m-%d}")

if __name__ == "__main__":
    main()
