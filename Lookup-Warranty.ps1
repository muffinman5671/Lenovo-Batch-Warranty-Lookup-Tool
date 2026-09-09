#Requires -Version 5.1
<#
.SYNOPSIS
    Command line batch warranty lookup for Lenovo serial numbers.

.DESCRIPTION
    The scriptable counterpart to the GUI. Reads serials from arguments, a text
    file, or the pipeline, and writes the main device warranty end date for each
    one in the order supplied. Battery and other component coverage is ignored.

.PARAMETER SerialNumber
    Serial numbers to look up.

.PARAMETER Path
    A text or CSV file of serial numbers, one per line.

.PARAMETER DateFormat
    Layout for the dates. Defaults to MM/dd/yyyy.

.PARAMETER DatesOnly
    Emit just the dates, one per line, ready to paste into a single Excel column.

.PARAMETER CsvPath
    Also write the full result set to this CSV file.

.EXAMPLE
    .\Lookup-Warranty.ps1 PF0ABCDE, PF1FGHIJ

.EXAMPLE
    .\Lookup-Warranty.ps1 -Path .\serials.txt -DatesOnly | Set-Clipboard

.EXAMPLE
    .\Lookup-Warranty.ps1 -Path .\serials.txt -CsvPath .\warranty.csv
#>
[CmdletBinding(DefaultParameterSetName = 'Args')]
param(
    [Parameter(ParameterSetName = 'Args', Position = 0, ValueFromPipeline)]
    [string[]] $SerialNumber,

    [Parameter(ParameterSetName = 'File', Mandatory)]
    [string] $Path,

    [ValidateSet('MM/dd/yyyy', 'yyyy-MM-dd', 'dd/MM/yyyy', 'M/d/yyyy')]
    [string] $DateFormat = 'MM/dd/yyyy',

    [switch] $DatesOnly,

    [string] $CsvPath,

    [ValidateRange(1, 32)]
    [int] $ThrottleLimit = 8
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    Import-Module (Join-Path $PSScriptRoot 'LenovoWarranty.psm1') -Force
    $collected = New-Object System.Collections.Generic.List[string]
}

process {
    if ($PSCmdlet.ParameterSetName -eq 'Args' -and $SerialNumber) {
        foreach ($s in $SerialNumber) { [void]$collected.Add($s) }
    }
}

end {
    if ($PSCmdlet.ParameterSetName -eq 'File') {
        if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
        $raw = Get-Content -LiteralPath $Path -Raw
    } else {
        $raw = ($collected -join "`n")
    }

    $serials = @(ConvertTo-SerialList -Text $raw)
    if ($serials.Count -eq 0) { throw 'No serial numbers supplied.' }

    Write-Verbose "Looking up $($serials.Count) serial number(s)."

    $shown = 0
    $tick  = {
        param($doneCount, $totalCount)
        Write-Progress -Activity 'Lenovo warranty lookup' `
                       -Status "$doneCount of $totalCount" `
                       -PercentComplete (100 * $doneCount / $totalCount)
    }

    $results = @(Get-LenovoWarranty -SerialNumber $serials -ThrottleLimit $ThrottleLimit -OnProgress $tick)
    Write-Progress -Activity 'Lenovo warranty lookup' -Completed

    if ($CsvPath) {
        $results | ForEach-Object {
            [pscustomobject]@{
                Serial      = $_.Serial
                WarrantyEnd = Format-WarrantyDate -Date $_.WarrantyEnd -Format $DateFormat
                Status      = $_.Status
                Product     = $_.Product
                MachineType = $_.MachineType
                Coverage    = $_.WarrantyName
                Note        = $_.Error
            }
        } | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "Saved $($results.Count) rows to $CsvPath" -ForegroundColor Green
    }

    if ($DatesOnly) {
        # Bare column: one line per input serial, blanks where nothing was found,
        # so it pastes into Excel alongside the original list without shifting.
        $results | ForEach-Object { Format-WarrantyDate -Date $_.WarrantyEnd -Format $DateFormat }
        return
    }

    $results | Select-Object `
        Serial,
        @{ Name = 'WarrantyEnd'; Expression = { Format-WarrantyDate -Date $_.WarrantyEnd -Format $DateFormat } },
        Status,
        Product,
        @{ Name = 'Coverage'; Expression = { if ($_.Error) { $_.Error } else { $_.WarrantyName } } }
}
