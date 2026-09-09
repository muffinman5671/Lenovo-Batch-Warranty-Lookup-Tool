#Requires -Version 5.1
<#
    LenovoWarranty.psm1

    Core lookup engine for the Lenovo Batch Warranty Lookup Tool.

    Talks to the same public endpoint the Lenovo support site itself uses when
    you open a product's warranty page:

        POST https://pcsupport.lenovo.com/us/en/api/v4/upsell/redport/getIbaseInfo
        {"serialNumber":"PF0ABCDE"}

    No API key, cookie or login is required.

    The response splits coverage into category "MACHINE" (the device itself)
    and category "COMPONENT" (battery, pen, sealed battery, and so on). This
    module reports ONLY the MACHINE coverage, which is the main device warranty
    people actually want, and ignores the battery/component entries.
#>

Set-StrictMode -Version Latest

$script:ApiUrl    = 'https://pcsupport.lenovo.com/us/en/api/v4/upsell/redport/getIbaseInfo'
$script:UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36'

# Lenovo is TLS 1.2+ only; Windows PowerShell 5.1 still defaults lower on some builds.
try {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch { }

Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue


function ConvertTo-DateOrNull {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }

    $parsed = [datetime]::MinValue
    $ok = [datetime]::TryParseExact(
            $Text.Trim(), 'yyyy-MM-dd',
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None, [ref]$parsed)
    if ($ok) { return $parsed }

    # Lenovo occasionally returns a full timestamp instead of a bare date.
    if ([datetime]::TryParse($Text, [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $parsed }

    return $null
}


function New-LenovoHttpClient {
    <#  Builds an HttpClient configured the way the support site calls the API. #>
    [CmdletBinding()]
    param([int]$TimeoutSec = 40)

    $handler = New-Object System.Net.Http.HttpClientHandler
    if ($handler.SupportsAutomaticDecompression) {
        $handler.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor
                                          [System.Net.DecompressionMethods]::Deflate
    }

    $client = New-Object System.Net.Http.HttpClient($handler)
    $client.Timeout = [TimeSpan]::FromSeconds($TimeoutSec)
    $client.DefaultRequestHeaders.Add('User-Agent', $script:UserAgent)
    $client.DefaultRequestHeaders.Add('Accept', 'application/json, text/plain, */*')

    return $client
}


function ConvertFrom-LenovoIbaseInfo {
    <#
        Turns one raw getIbaseInfo payload into the single flat result row the
        tool reports. Records problems in the Error field rather than throwing,
        so one bad serial never derails a batch.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Serial,
        [string] $Json,
        [string] $TransportError
    )

    $row = [ordered]@{
        Serial        = $Serial
        WarrantyEnd   = $null      # main device warranty expiry
        WarrantyStart = $null
        WarrantyName  = ''
        Product       = ''
        MachineType   = ''
        ShipDate      = $null
        Status        = ''         # In warranty / Out of warranty
        Error         = ''
    }

    if ($TransportError) {
        $row.Error = $TransportError
        return [pscustomobject]$row
    }

    try {
        $parsed = $Json | ConvertFrom-Json
    } catch {
        $row.Error = 'Unreadable response from Lenovo'
        return [pscustomobject]$row
    }

    # code 0 = success. Anything else carries a human readable reason,
    # e.g. code 100 "Machine Type and Serial Number Not Found in IBase."
    if ($parsed.code -ne 0) {
        $desc = ''
        if ($parsed.PSObject.Properties['msg'] -and $parsed.msg) { $desc = [string]$parsed.msg.desc }
        if (-not $desc) { $desc = "Lenovo returned code $($parsed.code)" }

        # Tidy Lenovo's internal wording so the results grid stays readable.
        $desc = $desc -replace '^Call sde api:\s*', ''
        switch -Regex ($desc) {
            'Not Found in IBase' { $desc = 'Serial number not found' }
            '^SN invalid'        { $desc = 'Invalid serial number' }
        }
        $row.Error = $desc.Trim()
        return [pscustomobject]$row
    }

    if (-not $parsed.PSObject.Properties['data'] -or -not $parsed.data) {
        $row.Error = 'No warranty data returned'
        return [pscustomobject]$row
    }
    $data = $parsed.data

    if ($data.PSObject.Properties['machineInfo'] -and $data.machineInfo) {
        $row.Product     = [string]$data.machineInfo.productName
        $row.MachineType = [string]$data.machineInfo.product
        $row.ShipDate    = ConvertTo-DateOrNull ([string]$data.machineInfo.shipDate)
    }
    if ($data.PSObject.Properties['warrantyStatus']) {
        $row.Status = [string]$data.warrantyStatus
    }

    # Gather every warranty entry Lenovo lists for this machine, then keep only
    # the ones covering the device itself. COMPONENT rows are battery/pen/etc.
    $all = @()
    foreach ($bucket in 'baseWarranties', 'upgradeWarranties', 'contractWarranties') {
        if ($data.PSObject.Properties[$bucket] -and $data.$bucket) { $all += @($data.$bucket) }
    }
    $machine = @($all | Where-Object { $_ -and $_.category -eq 'MACHINE' })

    if ($machine.Count -eq 0) {
        # Fall back to whatever Lenovo flags as the active warranty, provided it
        # is device level. A few old records carry no categorised buckets.
        if ($data.PSObject.Properties['currentWarranty'] -and $data.currentWarranty -and
            $data.currentWarranty.category -eq 'MACHINE') {
            $machine = @($data.currentWarranty)
        } else {
            $row.Error = 'No device warranty on record'
            return [pscustomobject]$row
        }
    }

    # The device is covered until the LAST machine level entry expires, which
    # correctly accounts for upgrades and extended contracts stacked on the base.
    $best     = $null
    $bestDate = $null
    foreach ($w in $machine) {
        $d = ConvertTo-DateOrNull ([string]$w.endDate)
        if ($null -eq $d) { continue }

        if ($null -eq $bestDate -or $d -gt $bestDate) {
            $bestDate = $d
            $best     = $w
        }
        elseif ($d -eq $bestDate -and $best.type -eq 'BASE' -and $w.type -ne 'BASE') {
            # Same expiry, but an upgrade or contract describes the real level of
            # service better than the base entry it sits on top of.
            $best = $w
        }
    }

    if ($null -eq $bestDate) {
        $row.Error = 'Device warranty has no end date'
        return [pscustomobject]$row
    }

    $row.WarrantyEnd   = $bestDate
    $row.WarrantyStart = ConvertTo-DateOrNull ([string]$best.startDate)
    $row.WarrantyName  = [string]$best.name

    return [pscustomobject]$row
}


function ConvertTo-SerialList {
    <#
        Accepts whatever the user pasted - one per line, comma separated, or a
        column copied straight out of Excel - and returns a clean ordered list.
        Order is preserved exactly, and duplicates are kept so the output lines
        up row for row with the input.
    #>
    [CmdletBinding()]
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }

    $out = New-Object System.Collections.Generic.List[string]
    foreach ($token in ($Text -split '[\r\n,;\t ]+')) {
        $t = $token.Trim().Trim('"').Trim()
        if ($t) { [void]$out.Add($t.ToUpperInvariant()) }
    }

    return $out.ToArray()
}


function Get-LenovoWarranty {
    <#
    .SYNOPSIS
        Looks up the main device warranty end date for a batch of Lenovo serials.

    .DESCRIPTION
        Queries Lenovo's public warranty API for each serial and returns one
        result object per serial, in exactly the order supplied. Only the
        device level ("MACHINE") warranty is reported; battery and other
        component coverage is ignored.

    .PARAMETER SerialNumber
        One or more Lenovo serial numbers.

    .PARAMETER ThrottleLimit
        How many requests to keep in flight at once. Default 8.

    .PARAMETER OnProgress
        Optional scriptblock invoked as the batch advances, receiving
        (completedCount, totalCount).

    .EXAMPLE
        Get-LenovoWarranty PF0ABCDE, PF1FGHIJ | Format-Table
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]] $SerialNumber,

        [ValidateRange(1, 32)]
        [int] $ThrottleLimit = 8,

        [int] $TimeoutSec = 40,

        [scriptblock] $OnProgress
    )

    begin { $queue = New-Object System.Collections.Generic.List[string] }

    process {
        foreach ($s in $SerialNumber) {
            if ($s -and $s.Trim()) { [void]$queue.Add($s.Trim()) }
        }
    }

    end {
        $serials = $queue.ToArray()
        $total   = $serials.Count
        if ($total -eq 0) { return }

        $client  = New-LenovoHttpClient -TimeoutSec $TimeoutSec
        $results = New-Object 'object[]' $total
        $apiUrl  = $script:ApiUrl
        $done    = [ref] 0

        # Runs one pass over the supplied indices, in fixed size waves. Results
        # are written back by index, so the output order always matches the
        # input order no matter which request happens to finish first.
        $runPass = {
            param([int[]] $Indices, [bool] $CountProgress)

            for ($offset = 0; $offset -lt $Indices.Count; $offset += $ThrottleLimit) {
                $count   = [Math]::Min($ThrottleLimit, $Indices.Count - $offset)
                $pending = @()

                for ($i = 0; $i -lt $count; $i++) {
                    $idx  = $Indices[$offset + $i]
                    $body = '{"serialNumber":"' + ($serials[$idx] -replace '["\\]', '') + '"}'
                    $content = New-Object System.Net.Http.StringContent(
                                    $body, [Text.Encoding]::UTF8, 'application/json')
                    $pending += [pscustomobject]@{
                        Index = $idx
                        Task  = $client.PostAsync($apiUrl, $content)
                    }
                }

                foreach ($p in $pending) {
                    $json = $null
                    $err  = $null
                    try {
                        $resp = $p.Task.GetAwaiter().GetResult()
                        if ($resp.IsSuccessStatusCode) {
                            $json = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        } else {
                            $err = "HTTP $([int]$resp.StatusCode) from Lenovo"
                        }
                        $resp.Dispose()
                    } catch {
                        $inner = $_.Exception
                        while ($inner.InnerException) { $inner = $inner.InnerException }
                        $err = "Request failed: $($inner.Message)"
                    }

                    $results[$p.Index] = ConvertFrom-LenovoIbaseInfo `
                                            -Serial $serials[$p.Index] `
                                            -Json $json -TransportError $err

                    if ($CountProgress) {
                        $done.Value++
                        if ($OnProgress) { & $OnProgress $done.Value $total }
                    }
                }
            }
        }

        try {
            & $runPass @(0..($total - 1)) $true

            # A throttled or briefly unavailable request is worth another try on a
            # large batch. A "serial not found" answer is final and is left alone.
            for ($attempt = 1; $attempt -le 2; $attempt++) {
                $retry = @(
                    0..($total - 1) | Where-Object {
                        $results[$_] -and $results[$_].Error -match 'HTTP (429|5\d\d)|Request failed'
                    }
                )
                if ($retry.Count -eq 0) { break }

                Start-Sleep -Milliseconds (400 * $attempt)
                & $runPass $retry $false
            }
        } finally {
            $client.Dispose()
        }

        return $results
    }
}


function Format-WarrantyDate {
    <#  Renders a warranty date in the layout the user picked for Excel. #>
    [CmdletBinding()]
    param(
        $Date,

        [ValidateSet('MM/dd/yyyy', 'yyyy-MM-dd', 'dd/MM/yyyy', 'M/d/yyyy')]
        [string] $Format = 'MM/dd/yyyy'
    )

    if ($null -eq $Date) { return '' }
    if ($Date -isnot [datetime]) { return '' }
    if ($Date -eq [datetime]::MinValue) { return '' }

    return $Date.ToString($Format, [Globalization.CultureInfo]::InvariantCulture)
}


Export-ModuleMember -Function Get-LenovoWarranty, ConvertTo-SerialList,
                              Format-WarrantyDate, ConvertFrom-LenovoIbaseInfo,
                              ConvertTo-DateOrNull
