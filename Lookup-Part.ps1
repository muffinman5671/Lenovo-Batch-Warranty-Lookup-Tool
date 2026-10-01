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

.EXAMPLE
    .\Lookup-Part.ps1 PF0ABCDE 'System board'

.EXAMPLE
    .\Lookup-Part.ps1 PF0ABCDE -Part SSD -PartNumbersOnly | Set-Clipboard

.EXAMPLE
    .\Lookup-Part.ps1 PF0ABCDE -CsvPath .\parts.csv

.EXAMPLE
    .\Lookup-Part.ps1 -ListParts
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
    [string] $CsvPath
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

if ($result.Error) { throw "$($serial[0]): $($result.Error)" }

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
