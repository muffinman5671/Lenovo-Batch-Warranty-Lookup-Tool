#Requires -Version 5.1
<#
.SYNOPSIS
    Command line part number lookup for one Lenovo serial number.

.DESCRIPTION
    The scriptable counterpart to the PARTS section of the GUI. Resolves the
    serial to its machine type and model, pulls Lenovo's parts list for it,
    and prints the part numbers under the commodity you name - Lenovo's own
    grouping for that machine - or the whole list.

.PARAMETER SerialNumber
    The Lenovo serial number.

.PARAMETER Commodity
    Which of Lenovo's commodities you want, as -ListCommodities prints them
    for that machine: "SYSTEM BOARDS", "LCD ASSEMBLIES", ... Case, spacing
    and punctuation do not matter, and a shorter form is fine when it fits
    only one ("system board"). Leave it out, or pass "All parts", for the
    whole list.

.PARAMETER ListCommodities
    Print the commodities Lenovo groups this machine's parts under, and exit.

.PARAMETER PartNumbersOnly
    Emit just the part numbers, one per line, ready for the clipboard.

.PARAMETER CsvPath
    Also write the matching rows to this CSV file.

.PARAMETER Diagnose
    Print what Lenovo answered to every attempt, then scan the product's
    parts page and its scripts for the parts API the site itself calls. The
    same report is saved next to this script as
    "Lenovo parts diagnostics <date>.txt", ready to paste into a bug report.

.EXAMPLE
    .\Lookup-Part.ps1 PF0ABCDE 'System boards'

.EXAMPLE
    .\Lookup-Part.ps1 PF0ABCDE -Commodity 'solid state drives' -PartNumbersOnly | Set-Clipboard

.EXAMPLE
    .\Lookup-Part.ps1 PF0ABCDE -CsvPath .\parts.csv

.EXAMPLE
    .\Lookup-Part.ps1 PF0ABCDE -ListCommodities

.EXAMPLE
    .\Lookup-Part.ps1 PF0ABCDE -Diagnose
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [string] $SerialNumber,

    [Parameter(Position = 1)]
    [Alias('Part')]
    [string] $Commodity = 'All parts',

    [switch] $ListCommodities,

    [switch] $PartNumbersOnly,

    [string] $CsvPath,

    [switch] $Diagnose
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'LenovoWarranty.psm1') -Force

$serial = @(ConvertTo-SerialList -Text $SerialNumber)
if ($serial.Count -ne 1) { throw 'Give exactly one serial number.' }

Write-Verbose "Looking up '$Commodity' for $($serial[0])."
$result = Find-LenovoPart -SerialNumber $serial[0] -Commodity $Commodity

if ($Diagnose) {
    $report     = @(Get-LenovoPartsDiagnostic -SerialNumber $result.Serial -PartsList $result)
    $reportPath = Join-Path $PSScriptRoot "Lenovo parts diagnostics $(Get-Date -Format 'yyyy-MM-dd').txt"
    $report | Set-Content -LiteralPath $reportPath -Encoding UTF8
    $report | ForEach-Object { Write-Host $_ }
    Write-Host "Saved to $reportPath" -ForegroundColor Green
    if ($result.Error) { return }
}

if ($result.Error) {
    # -Verbose shows what Lenovo actually sent back, which is what you need
    # when a serial you know is good comes back "not found".
    if ($result.Raw) { foreach ($l in ($result.Raw -split "`n")) { Write-Verbose "Lenovo replied: $l" } }
    throw "$($serial[0]): $($result.Error) (run with -Diagnose to see what Lenovo answered)"
}

Write-Verbose ("{0} - type {1}, model {2}: {3} parts listed, {4} matching" -f
    $result.Product, $result.MachineType, $result.Model, $result.Parts.Count, $result.Matches.Count)

if ($ListCommodities) {
    Write-Host "$($result.Product) - $($result.Parts.Count) parts under $($result.Commodities.Count) commodities:" -ForegroundColor Green
    'All parts'
    $result.Commodities
    return
}

if ($result.Note) {
    # The commodity asked for is not one this machine has; say which it has.
    throw "$($result.Serial): $($result.Note)"
}

$rows = @($result.Matches | ForEach-Object {
    [pscustomobject]@{
        Serial      = $result.Serial
        Product     = $result.Product
        MachineType = $result.MachineType
        PartNumber  = $_.PartNumber
        Description = $_.Description
        Commodity   = $_.Commodity
        Serviceable = $_.Cru
    }
})

if ($CsvPath) {
    $rows | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Saved $($rows.Count) rows to $CsvPath" -ForegroundColor Green
}

if ($rows.Count -eq 0) {
    Write-Warning "No parts listed for $($result.Serial) ($($result.Product))."
    return
}

if ($PartNumbersOnly) {
    $rows | ForEach-Object { $_.PartNumber }
    return
}

$rows | Select-Object PartNumber, Description, Commodity, Serviceable
