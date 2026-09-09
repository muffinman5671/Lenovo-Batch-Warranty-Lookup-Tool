# Lenovo Batch Warranty Lookup Tool

Paste a batch of Lenovo serial numbers, get back the **main device warranty end
date** for each one — in the same order you entered them, ready to paste
straight into Excel.

Built as a replacement for Lenovo's own
[batch warranty lookup](https://pcsupport.lenovo.com/us/en/warrantylookup/batchquery),
which requires building an upload file, returns results in its own order, and
mixes battery and other component coverage in with the device warranty.

## Quick start

Double-click **`Lenovo Warranty Lookup.cmd`**.

1. Paste your serials into the left box, one per line.
2. Click **Look up warranties** (or press `Ctrl+Enter`).
3. Click **Copy dates** and paste into your spreadsheet.

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

## Command line

For scripted use or very large batches:

```powershell
# Straight from arguments
.\Lookup-Warranty.ps1 PF0ABCDE, PF1FGHIJ

# From a file, dates only, onto the clipboard
.\Lookup-Warranty.ps1 -Path .\serials.txt -DatesOnly | Set-Clipboard

# Full result set to CSV
.\Lookup-Warranty.ps1 -Path .\serials.txt -CsvPath .\warranty.csv
```

The module can also be used directly:

```powershell
Import-Module .\LenovoWarranty.psm1
Get-LenovoWarranty PF0ABCDE, PF1FGHIJ | Format-Table
```

## Files

| File | Purpose |
|---|---|
| `Lenovo Warranty Lookup.cmd` | Double-click launcher for the GUI |
| `LenovoWarrantyLookup.ps1` | The GUI |
| `Lookup-Warranty.ps1` | Command line front end |
| `LenovoWarranty.psm1` | Lookup engine — all the API and parsing logic |

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

## Requirements

Windows PowerShell 5.1 (built into Windows) and internet access to
`pcsupport.lenovo.com`. No install, no dependencies.

## Notes and caveats

- This rides an undocumented internal endpoint. It has been stable for years,
  but if Lenovo changes it, `$script:ApiUrl` and `ConvertFrom-LenovoIbaseInfo`
  in `LenovoWarranty.psm1` are the two places to fix.
- Serials are normalised to uppercase; blank lines, commas, tabs and stray
  quotes in pasted input are handled.
- Duplicate serials are deliberately **kept**, so the output stays row-for-row
  with a column pasted out of a spreadsheet.
