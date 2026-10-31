---
name: p6-cost-loader
description: Load a Bill of Quantities (BOQ) into a Primavera P6 schedule as cost-loaded activity expenses - no typing BOQ lines into P6 by hand. Reads a BOQ workbook (or a folder of division files) and a P6 XER export, matches every BOQ line to the right activity with HIGH/MEDIUM/LOW confidence, splits it across the floors/zones where that activity exists (or onto the floor the BOQ names), checks the totals to the cent, flags BOQ errors (mis-filed lines, 10x rate typos, duplicates, lumped prices), and writes a P6-importable XER (original schedule + expenses with cost, quantity and unit), a P6 spreadsheet-import file and a review workbook. Weak matches never load money - they are listed for the planner. Use when a planner/QS/cost controller wants to cost-load or resource-load a P6 schedule from a BOQ, map BOQ items to activities, link BOQ and schedule, build a BOQ-activity mapping sheet, or check a cost-loaded schedule against the BOQ.
---

# P6 Cost Loader (BOQ -> Primavera P6)

Planners cost-load a P6 schedule by copying BOQ lines into activities by hand: which activity does
this line belong to, which floor, what share. Hundreds of lines, redone every BOQ revision, and every
typo is a wrong cash-flow curve. This skill does the transfer and shows its working.

## What it needs
- **BOQ**: an .xlsx workbook, or a folder of them (one file per CSI division is fine). Header row
  with Description + Qty + Rate/Amount; headings can be a column or text rows above the lines.
- **Schedule**: a P6 XER export of ONE project (File > Export > Primavera PM (XER)).
- Python 3 + `openpyxl`.

## Workflow (always two steps - a person approves the matches before anything is loaded)
1. **Match**
   ```
   py scripts/boq_to_p6.py match <boq.xlsx or folder> <schedule.xer> --out outputs/<job>
   ```
   Writes `<job>-mapping.csv`: every BOQ line, the activity it matched, split rule, confidence, flags.
2. **Review** - open the mapping (Excel is fine). Change `activity_group` to an activity name (all its
   floors) or a single Activity ID; set `split` to `equal` or `Level 01:40;Level 02:60`. Fix every LOW.
3. **Build**
   ```
   py scripts/boq_to_p6.py build <boq> <schedule.xer> --out outputs/<job>
   ```
   Writes `<job>.xer` (import: File > Import > Primavera PM (XER) > Create New Project, or Update
   Existing), `<job>-p6-import.xlsx`, `<job>-review.xlsx`.
   Build refuses a mapping made for a different/changed BOQ.

`run` = match + build in one go with no review (for testing only).
`--expenses line|heading|auto`: one P6 expense per BOQ line, or per BOQ heading per activity
(auto = per heading above 300 lines, so P6 gets hundreds of expenses, not tens of thousands).

## BOQ revisions (keep approved decisions, review only what changed)
```
py scripts/boq_to_p6.py revise <revised boq> <schedule.xer> --previous outputs/<job>-mapping.csv --out outputs/<job>-revB
```
Lines are paired by heading + description, not row number, so inserted lines shift nothing.
Writes `<job>-revB-mapping.csv` (with `revision_status`) and `<job>-revB-revision.xlsx` (reconciled
totals, To Review biggest money first, Deleted). Then review and run `build` with the same `--out`.

| Status | What happens |
|---|---|
| UNCHANGED | previous activity + split carried, no review |
| CHANGED | same line, new amount: decision carried at the new amount, listed for review |
| REWORDED | small wording change or moved heading: matched fresh, old decision shown beside it |
| NEW | matched fresh |
| SCHEDULE_CHANGED | the activity or a split zone left the schedule: not carried as it was |
| UNASSIGNED | still no activity from last time |

The previous mapping is trusted as approved, so pass the file the planner signed off. Build writes the
full expense set for the revised BOQ, not a delta: import as a new project, or clear the old expenses first.

## How to use this as Claude
- Run `match` first. Read the review's **Summary** and **Heading Review** (LOW first, biggest money
  first) back to the user: totals, how many lines are HIGH/MEDIUM/LOW, and every BOQ error flagged.
- Help the user fix LOWs in the mapping CSV - suggest the activity using the weak guess and the BOQ
  description, but let them decide. Never load a LOW line on your own judgment without saying so.
- Then run `build`, confirm "Totals check: PASS" (or explain exactly which lines are unassigned),
  and give the P6 import steps.
- If the schedule has no activity for a BOQ system (e.g. medical gases, UPS), say so - the fix is a
  new activity in P6, not forcing the money onto a near name.

## Confidence
| | Meaning |
|---|---|
| HIGH | the heading's words match the activity's name; no flags |
| MEDIUM | partial match, one-word match, cross-trade match, or a flag - glance at it |
| LOW | no clear activity - **not loaded**; weak guess shown |

Context used: CSI division of the line (`config/csi_divisions.json`), trade synonyms on the BOQ side
(`config/synonyms.json`), rare words weigh more than common ones, BOQ boilerplate ("supply,
installation and testing of") ignored, floor headings ("4TH FLOOR") route a line to that floor.

## Tests
`py tests/regression.py` - 79 checks: a planner-made case study (48/48 activities equal the planner's
manual loading), a real 1,332-line tower BOQ on a 688-activity schedule, floor routing, Excel-edited
mappings, BOQ revisions. `py tests/boq_reader_check.py` - BOQ reader vs a planner's cleaned BOQ (442/444 lines,
total diff 0.00). Test data is not bundled - see README. Run both after any change.

## Limits
Text/number BOQs only (no scanned PDFs). Does not create activities, logic or durations - the schedule
must exist. Matching is word-based: a schedule without an activity for some work leaves that money
unassigned (by design). Lump-sum lines split equally unless you set a split.
