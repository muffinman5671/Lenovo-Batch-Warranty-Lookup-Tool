#Requires -Version 5.1
<#
.SYNOPSIS
    Command line part number lookup for one Lenovo serial number.

.DESCRIPTION
    The scriptable counterpart to the PARTS section of the GUI. Resolves the
    serial to its machine type and model, pulls Lenovo's parts list for it,
    and prints the part number(s) for the part you name - or the whole list.

.PARAMETER SerialNumber
    The Lenovo serial number.

.PARAMETER Part
    Which part you want, from the preset list (see -ListParts). Case and
    punctuation do not matter: "system board", "LCD back cover", "ssd".
    Leave it out, or pass "All parts", for the whole list.

.PARAMETER ListParts
    Print the preset part names and exit.

.PARAMETER PartNumbersOnly
    Emit just the part numbers, one per line, ready for the clipboard.

.PARAMETER CsvPath
    Also write the matching rows to this CSV file.

.PARAMETER Diagnose
    When a lookup fails: print what Lenovo answered to every attempt, then
    scan the product's parts page and its scripts for the parts API the site
    itself calls. The same report is saved next to this script as
    "Lenovo parts diagnostics <date>.txt", ready to paste into a bug report.

.EXAMPLE
    .\Lookup-Part.ps1 PF0ABCDE 'System board'

.EXAMPLE
    .\Lookup-Part.ps1 PF0ABCDE -Part SSD -PartNumbersOnly | Set-Clipboard

.EXAMPLE
    .\Lookup-Part.ps1 PF0ABCDE -CsvPath .\parts.csv

.EXAMPLE
    .\Lookup-Part.ps1 -ListParts

.EXAMPLE
    .\Lookup-Part.ps1 PF0ABCDE -Diagnose
#>
[CmdletBinding(DefaultParameterSetName = 'Lookup')]
param(
    [Parameter(ParameterSetName = 'Lookup', Mandatory, Position = 0)]
    [string] $SerialNumber,

    [Parameter(ParameterSetName = 'Lookup', Position = 1)]
    [string] $Part = 'All parts',

    [Parameter(ParameterSetName = 'List', Mandatory)]
    [switch] $ListParts,

    [Parameter(ParameterSetName = 'Lookup')]
    [switch] $PartNumbersOnly,

    [Parameter(ParameterSetName = 'Lookup')]
    [string] $CsvPath,

    [Parameter(ParameterSetName = 'Lookup')]
    [switch] $Diagnose
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'LenovoWarranty.psm1') -Force

if ($PSCmdlet.ParameterSetName -eq 'List') {
    'All parts'
    Get-LenovoPartCategory | ForEach-Object { $_.Name }
    return
}

$serial = @(ConvertTo-SerialList -Text $SerialNumber)
if ($serial.Count -ne 1) { throw 'Give exactly one serial number.' }

Write-Verbose "Looking up '$Part' for $($serial[0])."
$result = Find-LenovoPart -SerialNumber $serial[0] -Part $Part

if ($Diagnose) {
    $report = New-Object System.Collections.Generic.List[string]
    [void]$report.Add("Lenovo part lookup diagnostics - $(Get-Date -Format 'yyyy-MM-dd HH:mm') - PowerShell $($PSVersionTable.PSVersion)")
    [void]$report.Add("Serial: $($result.Serial)")
    if ($result.Error) { [void]$report.Add("Result: $($result.Error)") }
    else { [void]$report.Add("Result: OK - $($result.Parts.Count) parts via $($result.Source) for $($result.Product)") }
    if ($result.Raw) {
        [void]$report.Add('What each attempt got back:')
        foreach ($l in ($result.Raw -split "`n")) { [void]$report.Add("  $l") }
    }
    [void]$report.Add('Scan of the parts page and its scripts:')
    foreach ($l in (Find-LenovoPartsEndpoint -SerialNumber $result.Serial)) { [void]$report.Add("  $l") }

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

$rows = @($result.Matches | ForEach-Object {
    [pscustomobject]@{
        Serial      = $result.Serial
        Product     = $result.Product
        MachineType = $result.MachineType
        PartNumber  = $_.PartNumber
        Description = $_.Description
        Commodity   = $_.Commodity
        Status      = $_.Status
        Substitutes = $_.Substitutes
        Cru         = $_.Cru
    }
})

if ($CsvPath) {
    $rows | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Saved $($rows.Count) rows to $CsvPath" -ForegroundColor Green
}

if ($rows.Count -eq 0) {
    Write-Warning "No '$Part' listed for $($result.Serial) ($($result.Product)). Try -Part 'All parts' to see the full list of $($result.Parts.Count) parts."
    return
}

if ($PartNumbersOnly) {
    $rows | ForEach-Object { $_.PartNumber }
    return
}

$rows | Select-Object PartNumber, Description, Commodity, Status, Substitutes
