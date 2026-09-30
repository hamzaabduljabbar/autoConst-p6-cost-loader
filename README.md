# P6 Cost Loader

**Load a Bill of Quantities into Primavera P6 without typing it in. Every BOQ line matched to its activity, split across the right floors, totals checked to the cent - and a P6 file you import in one click.**

Cost-loading a P6 schedule from a BOQ is usually manual: decide which activity each BOQ line belongs to, which floors it covers, what share each gets, then type it into P6. Hundreds of lines, redone every revision. This tool does the matching and the transfer, shows its reasoning for every line, and never loads money it isn't sure about.

Same honesty rule as the other AutoConst tools ([Drawing Takeoff](https://github.com/hamzaabduljabbar/autoConst-drawing-takeoff-claude), [Spec Bid-Review](https://github.com/hamzaabduljabbar/autoConst-spec-bid-review-claude)) - every number is traceable, nothing is guessed silently.

---

## Real results

**1. Planner case study** (planningengineer.net, a real planner's BOQ + P6 baseline + their manual cost loading):

```
BOQ total 840,414.00 | loaded 840,414.00 | diff 0.00
48/48 activities match the planner's manual cost loading
Imported into Primavera P6 Professional 25:  BOQ 1 Excavation  168 m3 x 86.13 = 14,469.00  (correct activity)
```
And it caught 8 BOQ errors the manual process pushed straight into P6: external paint filed under *Doors*, kitchen cabinets under *Waterproofing*, handrails under *wall tiles*, two skirting rates **10x apart**, a likely duplicate line.

**2. Real 17-floor tower BOQ** - 22 CSI division files, **1,332 priced lines, 106,996,378.52** (SAR):

| | |
|---|---|
| BOQ read vs the BOQ's own summary page | all 21 divisions equal, to the cent |
| BOQ read vs a planner's hand-cleaned copy | 442/444 lines, total diff 0.00, 439/442 headings agree |
| Run time | ~45 seconds |
| Auto-loaded onto a 688-activity schedule | **91%** (97.8M) |
| Left for the planner | 131 lines / 9.2M - shown with a weak guess, **never loaded blind** |
| What the planner reviews | 211 heading decisions (not 1,332 lines), LOW + biggest money first |
| Imported into P6 25 | 688 activities, 1,540 expenses, no errors |

Honest note: no schedule was published for this tower, so the 688-activity schedule was built for the test (trade x floor, names written without looking at the BOQ). In our own audit of the 211 decisions, about 2% of the auto-loaded value went to a near-but-wrong activity (all MEDIUM, all in the review list) - mostly systems the test schedule has no activity for (medical gases, UPS, motor control centre). **Review the Heading Review sheet before you import.**

---

## What you get

| File | What it is |
|---|---|
| `<job>-mapping.csv` | Every BOQ line -> activity, split, confidence, flags. **You edit this** (Excel is fine) |
| `<job>.xer` | Your original schedule, byte for byte, + a P6 expense per line (or per heading) with cost, quantity, unit |
| `<job>-p6-import.xlsx` | The same expenses in P6's spreadsheet-import format |
| `<job>-review.xlsx` | Summary, **Heading Review**, Matches, Activity Costs, Flags, Method |

## How it stays honest

- **A weak match never loads money.** LOW lines stay unassigned with a suggestion; wrong money on the wrong activity is worse than money waiting for you.
- **Totals check to the cent.** PASS only when every BOQ amount is loaded exactly once, line by line.
- **It never reads your answers.** Planner-added columns (activity, floor, %) are ignored.
- **BOQ errors are flagged, not fixed:** description vs heading conflicts, rate outliers, duplicates, a price with no description (lumped items).
- **Floor-aware:** a line under "4TH FLOOR" loads onto floor 4 only; otherwise it splits across the floors where that activity exists - and you can set any split.
- **Your schedule is untouched** - expenses are added, nothing else changes. Build refuses a mapping made for a different BOQ.

---

## Setting it up

### Step 1 - Install
**Windows:** `winget install --id Python.Python.3.12 -e` then `py -m pip install openpyxl`
**Mac:** `brew install python@3.12` then `python3 -m pip install openpyxl`

### Step 2 - Clone with Claude Code
> *"Clone https://github.com/hamzaabduljabbar/autoConst-p6-cost-loader and install it as a skill."*

### Step 3 - Export your schedule from P6
Open the project > **File > Export > Primavera PM (XER) > Project** > tick one project > save.

### Step 4 - Ask
> *"Cost-load my P6 schedule C:\jobs\tower.xer from the BOQ in C:\jobs\boq\"*

Claude runs the match, walks you through the LOW lines and the flags, then builds the file.

Or yourself:
```bash
py scripts/boq_to_p6.py match path/to/boq.xlsx path/to/schedule.xer --out outputs/job
# edit outputs/job-mapping.csv
py scripts/boq_to_p6.py build path/to/boq.xlsx path/to/schedule.xer --out outputs/job
```

### Step 5 - Import into P6
**File > Import > Primavera PM (XER) > Project** > choose `outputs/job.xer` > Import Action **Create New Project** (safest first time) > Finish. Check **Project > Expenses**.

---

## Tests

```bash
py tests/regression.py        # 65 checks: case study, 1,332-line tower, floor routing, Excel-edited mappings
py tests/boq_reader_check.py  # BOQ reader vs a planner's cleaned BOQ
```
Test data is **not bundled** (it belongs to planningengineer.net). Download, then place next to this repo:
- [How to create cost loading and BOQ mapping sheet to an existing schedule](https://planningengineer.net/download/how-to-create-cost-loading-and-boq-mapping-sheet-to-an-existing-schedule/) -> `../p6-test/hassan/` (`NEWPROJ-4.xer`, `Case Study BOQ.xlsx`)
- [How to read project BOQ - Inputs](https://planningengineer.net/download/how-to-read-project-boq-from-a-planning-point-of-view-inputs/) -> `../p6-test/hany/` (the 22 BOQ files)
- [How to read project BOQ - Outputs](https://planningengineer.net/download/how-to-read-project-boq-from-a-planning-point-of-view-outputs/) -> `../Data Loading.xlsx`

The tower schedule is generated by `tests/fixtures/make_tower_schedule.py`. Missing data = tests skip, never fake a pass.

## Folder layout
```
autoConst-p6-cost-loader/
  SKILL.md  README.md  LICENSE  requirements.txt
  scripts/boq_to_p6.py         <- the engine
  config/synonyms.json         <- trade wording (BOQ side)
  config/csi_divisions.json    <- CSI division -> trade words
  tests/regression.py  tests/boq_reader_check.py  tests/fixtures/make_tower_schedule.py
  outputs/                     <- your files (auto-created)
```

## Troubleshooting
- **"contains N projects"** - export ONE project to the XER.
- **"could not find a header row"** - the BOQ needs Description + Qty + Rate/Amount columns.
- **"mapping ... changed since 'match'"** - the BOQ changed; run `match` again.
- **"already has expenses"** - export the schedule without expenses, or remove them first.
- **Lots of LOW lines** - your schedule may have no activity for that work; add it in P6 or assign in the mapping.

## Limits
Excel BOQs only (no scanned PDFs). It loads cost onto an existing schedule - it does not create activities, logic or durations. Matching is word-based: odd activity names need a synonym (`config/synonyms.json`) or a mapping edit. Always review before a baseline.

---

Built by [Hamza Jabbar](https://hamzajabbar.online) - [AutoConst](https://hamzajabbar.online).
