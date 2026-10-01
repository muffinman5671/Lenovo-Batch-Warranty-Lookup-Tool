# Lenovo Batch Warranty Lookup Tool

Paste a batch of Lenovo serial numbers, get back the **main device warranty end
date** for each one — in the same order you entered them, ready to paste
straight into Excel.

A second section, **PARTS**, takes one serial number and returns the Lenovo
**part numbers** for that exact machine, narrowed to whichever of Lenovo's
own commodities (LCD ASSEMBLIES, SYSTEM BOARDS, ...) you pick.

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
machine?". It pulls the full FRU parts list Lenovo keeps for that serial,
then the **COMMODITY** dropdown fills with the groups Lenovo files that
machine's parts under — LCD ASSEMBLIES, SYSTEM BOARDS, CABLES INTERNAL,
SCREWS and so on, in Lenovo's own wording — and picking one narrows the grid
to it. The grouping differs per machine, which is why the list is built from
the lookup rather than fixed. **All parts** shows the whole list.

| Field | What it is |
|---|---|
| **Part no.** | The Lenovo FRU part number — what you order or quote |
| **Description** | Lenovo's own wording for the part |
| **Commodity** | The Lenovo group the part is filed under — handy on **All parts** |

A commodity can hold more than one part when a machine was sold in several
configurations — three LCD assemblies for touch, non-touch and WWAN, say —
so **read the description before ordering**.

Under the product card, **WARRANTY** shows when that machine's device
warranty ends and whether it is still active, with the coverage it comes
from — the same answer the warranty side gives, in the date format picked
there.

The parts list is downloaded once per serial; changing the commodity in the
dropdown re-filters it instantly.

- **Copy part no.** — the part numbers in the grid, one per line. Select
  rows first to copy just those.
- **Copy table** — serial, product, warranty end, part number, description
  and commodity, tab separated with a header row.

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
# Part numbers under one of Lenovo's commodities for one machine
.\Lookup-Part.ps1 PF0ABCDE 'System boards'

# Just the part numbers, onto the clipboard
.\Lookup-Part.ps1 PF0ABCDE -Commodity 'solid state drives' -PartNumbersOnly | Set-Clipboard

# The whole parts list for a serial, to CSV
.\Lookup-Part.ps1 PF0ABCDE -CsvPath .\parts.csv

# The commodities Lenovo groups this machine's parts under
.\Lookup-Part.ps1 PF0ABCDE -ListCommodities
```

The module can also be used directly:

```powershell
Import-Module .\LenovoWarranty.psm1
Get-LenovoWarranty PF0ABCDE, PF1FGHIJ | Format-Table

(Find-LenovoPart PF0ABCDE 'LCD assemblies').Matches
Get-LenovoPartCommodity (Get-LenovoPartsList PF0ABCDE).Parts
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
POST https://pcsupport.lenovo.com/us/en/api/v4/upsellAggregation/parts/asBuilt
     {"serialId":"pf0abcde","mtId":"21ah","model":"21ah00bbus"}
```

is the call the parts page itself makes when it shows "parts for your serial
number" — found by scanning the scripts the page loads. "As built" is the
parts list for that exact machine. The page also sends the CRU tiers it is
filtering on (self-service, optional-service, FRU), so the tool first asks
`parts/config` for the tier codes and sends those, falling back to the usual
codings. If as-built declines, the same body goes to `parts/model` (every part
for the model) and then `parts/compatible`, and the page's **Download parts
list** export is the last resort.

The JSON reader does not assume a layout: it walks whatever comes back and
keeps every object carrying a part number, so a bare list, a list under
`data`, or parts grouped under commodities all read the same, with the group's
commodity name carried onto each part. Lenovo's own as-built list is an
array of objects with the FRU number under `id`, the wording under `name`,
the family under `commodityVal` and the tier under `cruTier`; items with no
FRU (the "installed, but no further details" ones) have an empty `id` and
are left out. A spreadsheet export is read straight
out of the xlsx (it is only a zip of XML) — no Excel, no extra modules — and
CSV is handled too. Columns are matched by wording (`FRU`, `Part Number`,
`Description`, `Commodity`, ...) rather than position, and every column has
a list of names it answers to, so a row whose first-choice field is empty
falls back to the next one. The diagnostic report lists the field names the
reply actually used, with a sample of each.

## Requirements

Windows PowerShell 5.1 (built into Windows) and internet access to
`pcsupport.lenovo.com`. No install, no dependencies.

## Notes and caveats

- This rides an undocumented internal endpoint. It has been stable for years,
  but if Lenovo changes it, `$script:ApiUrl` and `ConvertFrom-LenovoIbaseInfo`
  in `LenovoWarranty.psm1` are the two places to fix.
- Part lookup rides two more of the same kind. The product resolver is widely
  used and well understood; the parts calls are what the site's own page
  scripts make, read out of those scripts rather than documented. If they
  change, `$script:PartsApiBase`, the attempt list in `Get-LenovoPartsList`,
  and the column patterns in `ConvertTo-LenovoPartRows` are the places to
  look, and the **Diagnose** report shows what the page calls now.
- The commodity dropdown is only as good as Lenovo's grouping. A part filed
  under an unexpected commodity is still in the list — **All parts** is
  always there as the backstop.
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
