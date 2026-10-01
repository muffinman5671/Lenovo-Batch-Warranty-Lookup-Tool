# Lenovo Batch Warranty Lookup Tool

Paste a batch of Lenovo serial numbers, get back the **main device warranty end
date** for each one — in the same order you entered them, ready to paste
straight into Excel.

A second section, **PARTS**, takes one serial number and a part picked from a
preset list (LCD back cover, SSD, system board, ...) and returns the Lenovo
**part number(s)** for that exact machine.

Built as a replacement for Lenovo's own
[batch warranty lookup](https://pcsupport.lenovo.com/us/en/warrantylookup/batchquery),
which requires building an upload file, returns results in its own order, and
mixes battery and other component coverage in with the device warranty.

## Quick start

Double-click **`Lenovo Warranty Lookup.cmd`**.

1. Paste your serials into the left box, one per line.
2. Click **Look up warranties** (or press `Ctrl+Enter`).
3. Click **Copy dates** and paste into your spreadsheet.

For a part number, click **PARTS** in the masthead:

1. Type the serial number.
2. Pick the part from the dropdown.
3. Press `Enter` (or click **Find part**). The part number(s) land in the
   grid; **Copy part no.** puts them on the clipboard.

## What it returns

Only the **device** warranty. Lenovo's API tags each warranty entry with a
category:

| Category | Example | Reported? |
|---|---|---|
| `MACHINE` | `3Y On-site, 9X5`, `3Y Depot 9X5 2BD` | **Yes** |
| `COMPONENT` | `1YR Battery`, `3Y Sealed Battery` | No |

Where a device has a base warranty plus an upgrade or an extended contract, the
tool reports the **latest** machine-level end date, which is when device
coverage actually runs out.

## Copy buttons

- **Copy dates** — one date per line, nothing else. Drops into a single Excel
  column lined up against the serials you pasted. A serial that could not be
  found leaves a **blank line**, so rows never shift out of alignment.
- **Copy table** — `Serial`, `Warranty End`, `Status`, `Product`, `Coverage`,
  tab separated with a header row.
- **Save CSV...** — the same columns plus machine type, written to a file.

Date format is selectable: `MM/dd/yyyy` (default), `yyyy-MM-dd`, `dd/MM/yyyy`,
`M/d/yyyy`. Changing it re-renders the results instantly — no need to look
anything up again.

## Part lookup

The **PARTS** section answers "what is the part number for the X on this
machine?". It pulls the full FRU parts list Lenovo keeps for that serial and
keeps the rows that match the part you picked.

| Field | What it is |
|---|---|
| **Part no.** | The Lenovo FRU part number — what you order or quote |
| **Description** | Lenovo's own wording for the part |
| **Commodity** | Lenovo's part family (covers, system boards, storage, ...) |
| **Status** | Availability as Lenovo reports it; unavailable parts show magenta |
| **Substitutes** | Replacement part numbers, when Lenovo lists any |

The preset list: LCD panel, LCD back cover, LCD bezel, LCD cable, hinges,
system board, power button board, I/O board, SSD, hard drive, memory,
battery, AC adapter, power cord, keyboard, palmrest / C cover, base cover /
D cover, touchpad, fingerprint reader, fan / heatsink, wireless card, antenna,
camera, speakers, screws — plus **All parts** for the whole list.

Matching goes by Lenovo's wording. Each preset has a pattern it looks for
(`PLANAR` and `SYSTEM BOARD` both count as a system board) and a pattern it
rules out (a `Bracket, SSD` is not an SSD). A part can match more than once
when a machine was sold in several configurations — three LCD back covers for
touch, non-touch and WWAN, say — so **read the description before ordering**.
If the preset you picked finds nothing, switch to **All parts** and scan.

The parts list is downloaded once per serial; changing the part in the
dropdown re-filters it instantly.

- **Copy part no.** — the part numbers in the grid, one per line. Select
  rows first to copy just those.
- **Copy table** — serial, product, part number, description, commodity,
  status and substitutes, tab separated with a header row.

## Look

Brutalist: hard edges, 3px rules, solid offset shadows on the action blocks,
monospace data, no gradients and no rounded corners. The palette is taken
straight from the project logo.

| Role | Colour | Used for |
|---|---|---|
| Crimson | `#E10B2E` | Primary action, warranty dates, active coverage block |
| Magenta | `#FF0F6B` | Accent, selection, problems |
| Navy | `#14142D` | Masthead, every rule and border, body text |
| Bone | `#F5F4F0` | Page, alternating rows |

Reading the grid: the **date** is crimson because it is the thing you came for,
a device still under coverage gets a solid `ACTIVE` block, and anything that
failed is magenta.

Created by **Aiden Ortega**.

## Command line

These are PowerShell scripts. From a PowerShell window run them as shown
below; from a plain Command Prompt, go through `powershell` instead:

```
powershell -NoProfile -ExecutionPolicy Bypass -File ".\Lookup-Part.ps1" PF0ABCDE -Diagnose
```

For scripted use or very large batches:

```powershell
# Straight from arguments
.\Lookup-Warranty.ps1 PF0ABCDE, PF1FGHIJ

# From a file, dates only, onto the clipboard
.\Lookup-Warranty.ps1 -Path .\serials.txt -DatesOnly | Set-Clipboard

# Full result set to CSV
.\Lookup-Warranty.ps1 -Path .\serials.txt -CsvPath .\warranty.csv
```

Part lookup has its own front end:

```powershell
# Part numbers for one part on one machine
.\Lookup-Part.ps1 PF0ABCDE 'System board'

# Just the part number(s), onto the clipboard
.\Lookup-Part.ps1 PF0ABCDE -Part SSD -PartNumbersOnly | Set-Clipboard

# The whole parts list for a serial, to CSV
.\Lookup-Part.ps1 PF0ABCDE -CsvPath .\parts.csv

# The preset part names
.\Lookup-Part.ps1 -ListParts
```

The module can also be used directly:

```powershell
Import-Module .\LenovoWarranty.psm1
Get-LenovoWarranty PF0ABCDE, PF1FGHIJ | Format-Table

(Find-LenovoPart PF0ABCDE 'LCD back cover').Matches
(Get-LenovoPartsList PF0ABCDE).Parts | Format-Table
```

## Files

| File | Purpose |
|---|---|
| `Lenovo Warranty Lookup.cmd` | Double-click launcher for the GUI |
| `Lenovo Parts Diagnostics.cmd` | Double-click launcher for the part lookup diagnostic |
| `LenovoWarrantyLookup.ps1` | The GUI |
| `Lookup-Warranty.ps1` | Command line front end for warranty lookup |
| `Lookup-Part.ps1` | Command line front end for part lookup |
| `LenovoWarranty.psm1` | Lookup engine — all the API and parsing logic for both |

## How it works

It calls the same public endpoint the Lenovo support site itself uses when you
open a product's warranty page:

```
POST https://pcsupport.lenovo.com/us/en/api/v4/upsell/redport/getIbaseInfo
{"serialNumber":"PF0ABCDE"}
```

No API key, cookie or login is needed. Because this is the per-device endpoint
rather than the batch queue, there is no upload step, no 1000-row cap and no
waiting on Lenovo's shared batch quota.

Requests run 8 at a time and results are written back **by index**, so output
order always matches input order regardless of which request finishes first.
Roughly 25 serials take about 4 seconds. Requests that come back throttled or
with a server error are retried twice; a "serial not found" answer is treated
as final.

### Part lookup

Two more calls to the same support site, chained:

```
GET https://pcsupport.lenovo.com/us/en/api/v4/mse/getproducts?productId=PF0ABCDE
```

resolves the serial to its product path, which carries the machine type and
model (`.../21ah/21ah00bbus/pf0abcde`), and then

```
POST https://pcsupport.lenovo.com/us/en/api/v4/upsellAggregation/parts/export
     ?type=SERIAL&serialId=pf0abcde&model=21ah00bbus&mtId=21ah
```

is the **Download parts list** link from Lenovo's own parts lookup page. It
returns the serial's full FRU list as a spreadsheet, which the module reads
straight out of the xlsx (it is only a zip of XML) — no Excel, no extra
modules. Should the export ever come back as CSV or JSON instead, the same
normaliser handles those too, matching columns by wording (`FRU`, `Part
Number`, `Description`, `Commodity`, ...) rather than position.

The export link was captured from a browser rather than documented, so the
exact request it wants is not certain. The tool tries the plausible shapes in
turn — POST with the query string, GET, POST with the same fields as a JSON
body, and POST again after loading the product's parts page so any cookies it
sets ride along — and takes the first that returns a parts list.

## Requirements

Windows PowerShell 5.1 (built into Windows) and internet access to
`pcsupport.lenovo.com`. No install, no dependencies.

## Notes and caveats

- This rides an undocumented internal endpoint. It has been stable for years,
  but if Lenovo changes it, `$script:ApiUrl` and `ConvertFrom-LenovoIbaseInfo`
  in `LenovoWarranty.psm1` are the two places to fix.
- Part lookup rides two more of the same kind. The product resolver is widely
  used and well understood; the parts export is the site's own download link,
  captured from a browser session rather than documented. If it changes,
  `$script:PartsExportUrl`, the query built in `Get-LenovoPartsList`, and the
  column patterns in `ConvertTo-LenovoPartRows` are the places to look. The
  preset patterns live in `$script:PartCategories` and are easy to extend.
- Part matching is textual. It is tuned to the wording Lenovo uses in its
  parts lists, but a part with an unusual description can be missed or an
  odd one included — **All parts** is always there as the backstop.
- When a part lookup fails, the product card in the GUI shows what Lenovo
  answered the first attempt. The **Diagnose** block (or
  `.\Lookup-Part.ps1 <serial> -Diagnose`, or double-clicking
  `Lenovo Parts Diagnostics.cmd`) writes a report with what every attempt
  got back, then scans the product's parts page and the scripts it loads
  for the parts API the site itself calls. The report lands next to the
  script as `Lenovo parts diagnostics <date>.txt`; the GUI opens it in
  Notepad. A valid serial coming back "not found" with a reply of `[]`
  means Lenovo's product resolver has no record of it; "Lenovo refused the
  parts list" means the export wants something the tool is not sending, and
  the scan is what shows what that is.
- Serials are normalised to uppercase; blank lines, commas, tabs and stray
  quotes in pasted input are handled.
- Duplicate serials are deliberately **kept**, so the output stays row-for-row
  with a column pasted out of a spreadsheet.
