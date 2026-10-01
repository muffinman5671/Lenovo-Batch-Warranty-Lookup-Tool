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




# ==========================================================================
#  Part lookup
#
#  Two more endpoints off the same support site, chained together:
#
#    1. GET  /us/en/api/v4/mse/getproducts?productId=PF0ABCDE
#       Resolves a serial to its product path, which carries the machine
#       type and model the parts list is keyed on:
#         laptops-and-netbooks/thinkpad-t-series-laptops/thinkpad-t14-gen-3-
#         type-21ah-21aj/21ah/21ah00bbus/pf0abcde
#
#    2. POST /us/en/api/v4/upsellAggregation/parts/asBuilt
#         {"serialId":"pf0abcde","mtId":"21ah","model":"21ah00bbus"}
#       The call the parts page itself makes (found by scanning the page's
#       scripts): the parts list for that serial as it was built. The same
#       body also works against parts/model and parts/compatible, which are
#       the fallbacks, and the page's "Download parts list" export is the
#       last resort.
#
#  The JSON reader walks whatever shape comes back and picks out every
#  object that carries a part number, inheriting a commodity name from the
#  group it sits under. A spreadsheet export is read straight out of the
#  xlsx (it is just a zip of XML) so there is no Excel dependency, and CSV
#  is handled too, so a change on Lenovo's side keeps working as long as
#  there is still a part number column.
# ==========================================================================

$script:SiteBase       = 'https://pcsupport.lenovo.com'
$script:ProductsUrl    = 'https://pcsupport.lenovo.com/us/en/api/v4/mse/getproducts'
$script:PartsApiBase   = 'https://pcsupport.lenovo.com/us/en/api/v4/upsellAggregation/parts'
$script:PartsExportUrl = 'https://pcsupport.lenovo.com/us/en/api/v4/upsellAggregation/parts/export'
$script:PartsReferer   = 'https://pcsupport.lenovo.com/us/en/partslookup'

# What a parts list header row looks like, used to skip any title lines
# Lenovo puts above it.
$script:PartsHeaderPattern = 'part\s*(number|no\b|no\.|#|num)|\bfru\b|descr'

# What a part number property is called in JSON, and what its value looks
# like (Lenovo FRU and option numbers mix letters and digits: 01AV430,
# 5CB1H89763, 4X40E77322). A purely numeric id is not one.
$script:PartNumberKeyPattern   = '(?i)fru|^pn$|pn$|^pn[_-]?|part_?(no|num|number|id|code)|material_?(no|num|number)|^sku$'
$script:PartNumberValuePattern = '^(?=.*[A-Za-z])(?=.*\d)[0-9A-Za-z-]{5,14}$'

Add-Type -AssemblyName System.IO.Compression            -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue


function Get-LenovoPartCommodity {
    <#
    .SYNOPSIS
        The commodities - Lenovo's own grouping of parts - present in a parts list.
    .DESCRIPTION
        The distinct commodity names in a parts list, in Lenovo's wording,
        sorted. This is what the GUI's dropdown fills with after a lookup:
        the list is per machine, so the choices are whatever Lenovo groups
        that machine's parts under. Parts without a commodity add nothing.
    #>
    [CmdletBinding()]
    param([Parameter(Position = 0)] [object[]] $Part)

    if ($null -eq $Part) { return @() }
    $seen  = @{}
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($p in $Part) {
        if ($null -eq $p -or -not $p.PSObject.Properties['Commodity']) { continue }
        $c = ([string]$p.Commodity).Trim()
        if (-not $c) { continue }
        $k = ($c -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
        if (-not $seen.ContainsKey($k)) { $seen[$k] = $true; [void]$names.Add($c) }
    }
    return @($names | Sort-Object)
}


function Resolve-LenovoPartCommodity {
    <#
        Matches what was asked for to one commodity of a parts list: exact
        once case, spacing and punctuation are ignored ("system boards" is
        SYSTEM BOARDS), else the one commodity that contains the words
        ("system board", "ssd" for SOLID STATE DRIVES would not). Throws,
        naming the machine's commodities, when nothing fits.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)] [object[]] $Part,
        [Parameter(Position = 1)] [string]   $Commodity
    )

    $flat  = { param($v) ([string]$v -replace '[^A-Za-z0-9]', '').ToUpperInvariant() }
    $names = @(Get-LenovoPartCommodity -Part $Part)
    $want  = & $flat $Commodity

    foreach ($n in $names) { if ((& $flat $n) -eq $want) { return $n } }
    if ($want) {
        $contains = @($names | Where-Object { (& $flat $_).Contains($want) })
        if ($contains.Count -eq 1) { return $contains[0] }
    }
    $have = if ($names.Count -gt 0) { $names -join ', ' } else { 'no commodities at all' }
    throw "No '$Commodity' commodity for this machine. It has: $have"
}


function Select-LenovoPart {
    <#
    .SYNOPSIS
        Filters a parts list down to one of Lenovo's commodities.
    .DESCRIPTION
        Keeps the rows whose commodity is the one named, matched the way
        Resolve-LenovoPartCommodity matches. An empty Commodity or 'All
        parts' returns the list untouched.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)] [object[]] $Part,
        [Parameter(Position = 1)] [string]   $Commodity
    )

    if ($null -eq $Part) { return @() }
    if (-not $Commodity -or $Commodity -match '^\s*all(\s*parts)?\s*$') { return @($Part) }

    $flat = { param($v) ([string]$v -replace '[^A-Za-z0-9]', '').ToUpperInvariant() }
    $name = & $flat (Resolve-LenovoPartCommodity -Part $Part -Commodity $Commodity)

    return @($Part | Where-Object { (& $flat $_.Commodity) -eq $name })
}


function ConvertFrom-LenovoJsonList {
    <#
        ConvertFrom-Json with the top level always handed back as an array of
        items. Windows PowerShell 5.1 emits a JSON array as ONE object (so
        wrapping it in @() nests it one level deep), while PowerShell 7
        unrolls it onto the pipeline. Either way this returns the elements.
        A JSON object comes back as a one element array.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Text)

    $parsed = @($Text | ConvertFrom-Json)
    while ($parsed.Count -eq 1 -and $null -ne $parsed[0] -and $parsed[0] -is [array]) {
        $parsed = @($parsed[0])
    }

    # Emitted element by element on purpose: callers wrap the call in @(),
    # which then yields a flat array on both 5.1 and 7.
    return $parsed
}


function Get-LenovoResponseSummary {
    <#  One line saying what Lenovo actually sent, for the attempt log. #>
    param($Response)

    $snippet = ''
    if ($Response.Text) {
        $snippet = ($Response.Text -replace '\s+', ' ').Trim()
        if ($snippet.Length -gt 400) { $snippet = $snippet.Substring(0, 400) + '...' }
    } elseif ($Response.Bytes -and $Response.Bytes.Length -gt 0) {
        $snippet = "<$($Response.Bytes.Length) bytes of binary>"
    }
    return "HTTP $($Response.StatusCode) $($Response.ContentType) $snippet".Trim()
}


function Invoke-LenovoRequest {
    <#
        One HTTP round trip, returned as a flat object rather than thrown, so
        callers can decide what a 404 or a 405 means for them.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Client,
        [Parameter(Mandatory)] [string] $Url,
        [ValidateSet('GET', 'POST')] [string] $Method = 'GET',
        [string] $Accept = 'application/json, text/plain, */*',
        [string] $Referer,
        [string] $Body,
        [string] $ContentType = 'application/json',
        [hashtable] $Headers
    )

    $out = [ordered]@{
        StatusCode  = 0
        ContentType = ''
        Bytes       = New-Object 'byte[]' 0     # never null, so .Length is always safe
        Text        = ''
        Error       = ''
    }

    try {
        $req = New-Object System.Net.Http.HttpRequestMessage(
                    [System.Net.Http.HttpMethod]::$Method, $Url)
        [void]$req.Headers.TryAddWithoutValidation('Accept', $Accept)
        if ($Referer) { $req.Headers.Referrer = [Uri]$Referer }
        if ($Headers) {
            foreach ($k in $Headers.Keys) { [void]$req.Headers.TryAddWithoutValidation($k, [string]$Headers[$k]) }
        }
        if ($Method -eq 'POST') {
            $payload = if ($null -ne $Body) { $Body } else { '' }
            $req.Content = New-Object System.Net.Http.StringContent(
                               $payload, [Text.Encoding]::UTF8, $ContentType)
        }

        $resp = $Client.SendAsync($req).GetAwaiter().GetResult()
        try {
            $out.StatusCode = [int]$resp.StatusCode
            if ($resp.Content) {
                if ($resp.Content.Headers.ContentType) {
                    $out.ContentType = [string]$resp.Content.Headers.ContentType.MediaType
                }
                $out.Bytes = $resp.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
            }
            if ($null -eq $out.Bytes) { $out.Bytes = New-Object 'byte[]' 0 }

            # Only decode as text when it is not a zip; a spreadsheet mangled
            # through UTF8 is useless and large.
            if ($out.Bytes.Length -lt 2 -or
                -not ($out.Bytes[0] -eq 0x50 -and $out.Bytes[1] -eq 0x4B)) {
                $out.Text = [Text.Encoding]::UTF8.GetString($out.Bytes)
            }
            if (-not $resp.IsSuccessStatusCode) {
                $out.Error = "HTTP $($out.StatusCode) from Lenovo"
            }
        } finally {
            $resp.Dispose()
        }
    } catch {
        $inner = $_.Exception
        while ($inner.InnerException) { $inner = $inner.InnerException }
        $out.Error = "Request failed: $($inner.Message)"
    }

    return [pscustomobject]$out
}


function Get-LenovoProduct {
    <#
    .SYNOPSIS
        Resolves a Lenovo serial number to its product, machine type and model.
    .EXAMPLE
        Get-LenovoProduct PF0ABCDE
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)] [string] $SerialNumber,
        [int] $TimeoutSec = 40,
        $Client
    )

    $serial = $SerialNumber.Trim().ToUpperInvariant()
    $row = [ordered]@{
        Serial      = $serial
        ProductId   = ''      # full lowercase product path from Lenovo
        Product     = ''      # readable product name
        MachineType = ''      # e.g. 21AH
        Model       = ''      # e.g. 21AH00BBUS
        Error       = ''
        Raw         = ''      # what Lenovo sent, kept when something goes wrong
    }

    if (-not $serial) {
        $row.Error = 'No serial number supplied'
        return [pscustomobject]$row
    }

    $ownClient = $false
    if (-not $Client) { $Client = New-LenovoHttpClient -TimeoutSec $TimeoutSec; $ownClient = $true }

    try {
        $url  = $script:ProductsUrl + '?productId=' + [Uri]::EscapeDataString($serial)
        $resp = Invoke-LenovoRequest -Client $Client -Url $url -Referer $script:PartsReferer

        $row.Raw = Get-LenovoResponseSummary $resp
        if ($resp.Error) {
            $row.Error = $resp.Error
            return [pscustomobject]$row
        }

        try {
            $parsed = @(ConvertFrom-LenovoJsonList -Text $resp.Text)
        } catch {
            $row.Error = 'Unreadable response from Lenovo'
            return [pscustomobject]$row
        }

        # Normally a one element array of products; be tolerant of a wrapper
        # object that carries the array under "data".
        if ($parsed.Count -eq 1 -and $null -ne $parsed[0] -and
            $parsed[0].PSObject.Properties['data'] -and
            -not ($parsed[0].PSObject.Properties['Id'] -or $parsed[0].PSObject.Properties['id'])) {
            $parsed = @($parsed[0].data)
            while ($parsed.Count -eq 1 -and $null -ne $parsed[0] -and $parsed[0] -is [array]) {
                $parsed = @($parsed[0])
            }
        }
        $entries = @($parsed | Where-Object {
            $_ -and ($_.PSObject.Properties['Id'] -or $_.PSObject.Properties['id'] -or $_.PSObject.Properties['ID'])
        })

        if ($entries.Count -eq 0) {
            $row.Error = 'Serial number not found'
            return [pscustomobject]$row
        }

        # A serial query comes back typed Product.Serial; prefer that entry
        # should Lenovo ever list the parent machine type alongside it.
        $entry = $entries | Where-Object {
            ($_.PSObject.Properties['Type']   -and [string]$_.Type   -eq 'Product.Serial') -or
            ($_.PSObject.Properties['Serial'] -and [string]$_.Serial -eq $serial)
        } | Select-Object -First 1
        if (-not $entry) { $entry = $entries[0] }

        $id = ''
        foreach ($k in 'Id', 'id', 'ID') {
            if ($entry.PSObject.Properties[$k] -and $entry.$k) { $id = [string]$entry.$k; break }
        }
        if (-not $id) {
            $row.Error = 'Serial number not found'
            return [pscustomobject]$row
        }

        $row.ProductId = $id.Trim('/').ToLowerInvariant()
        foreach ($k in 'Name', 'name', 'ProductName') {
            if ($entry.PSObject.Properties[$k] -and $entry.$k) { $row.Product = ([string]$entry.$k).Trim(); break }
        }

        # .../<machine type>/<model>/<serial>
        $seg = @($row.ProductId -split '/' | Where-Object { $_ })
        if ($seg.Count -ge 3 -and $seg[-1] -eq $serial.ToLowerInvariant()) {
            $row.Model       = $seg[-2].ToUpperInvariant()
            $row.MachineType = $seg[-3].ToUpperInvariant()
        } elseif ($seg.Count -ge 2) {
            # Lenovo answered with the model rather than the serial path.
            $row.Model       = $seg[-1].ToUpperInvariant()
            $row.MachineType = $seg[-2].ToUpperInvariant()
        }

        if (-not $row.MachineType) {
            $row.Error = 'Could not work out the machine type for this serial'
            return [pscustomobject]$row
        }
        $row.Raw = ''
        return [pscustomobject]$row
    } finally {
        if ($ownClient) { $Client.Dispose() }
    }
}


function ConvertFrom-LenovoXlsx {
    <#
        Reads the first worksheet of an xlsx straight out of the zip and
        returns one ordered dictionary per row, keyed by the header row. No
        Excel, no COM, no third party module.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [byte[]] $Bytes)

    $colIndex = {
        param([string] $Ref)
        $letters = ($Ref -replace '[^A-Za-z]', '').ToUpperInvariant()
        $n = 0
        foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
        return $n - 1
    }

    $readText = {
        # Rich text splits one string across several <t> runs; join them.
        param($Node)
        $sb = New-Object Text.StringBuilder
        foreach ($t in $Node.GetElementsByTagName('t')) { [void]$sb.Append($t.InnerText) }
        return $sb.ToString()
    }

    $stream = New-Object IO.MemoryStream(,$Bytes)
    $zip    = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Read)
    try {
        $shared = New-Object System.Collections.Generic.List[string]
        $ssEntry = $zip.GetEntry('xl/sharedStrings.xml')
        if ($ssEntry) {
            $reader = New-Object IO.StreamReader($ssEntry.Open())
            try { $ssXml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
            foreach ($si in $ssXml.DocumentElement.GetElementsByTagName('si')) {
                [void]$shared.Add((& $readText $si))
            }
        }

        $sheets = @($zip.Entries | Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' } |
                    Sort-Object { [int]($_.FullName -replace '\D', '') })
        if ($sheets.Count -eq 0) { throw 'The spreadsheet has no worksheets' }

        foreach ($sheet in $sheets) {
            $reader = New-Object IO.StreamReader($sheet.Open())
            try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }

            $rows = New-Object System.Collections.Generic.List[object]
            foreach ($r in $xml.DocumentElement.GetElementsByTagName('row')) {
                $cells = @{}
                foreach ($c in $r.GetElementsByTagName('c')) {
                    $ref  = $c.GetAttribute('r')
                    $type = $c.GetAttribute('t')
                    $val  = ''
                    if ($type -eq 'inlineStr') {
                        $val = & $readText $c
                    } else {
                        $vNode = $c.GetElementsByTagName('v')
                        if ($vNode.Count -gt 0) { $val = $vNode[0].InnerText }
                        if ($type -eq 's') {
                            $i = 0
                            if ([int]::TryParse($val, [ref]$i) -and $i -ge 0 -and $i -lt $shared.Count) {
                                $val = $shared[$i]
                            }
                        }
                    }
                    if ($ref) { $cells[(& $colIndex $ref)] = $val.Trim() }
                }
                if ($cells.Count -gt 0) { [void]$rows.Add($cells) }
            }

            # The header is the first row that talks about part numbers. Lenovo
            # may put a title line or the product name above it.
            $headerAt = -1
            for ($i = 0; $i -lt $rows.Count; $i++) {
                $texts = @($rows[$i].Values | Where-Object { $_ })
                if ($texts.Count -ge 2 -and ($texts -join ' ') -match $script:PartsHeaderPattern) { $headerAt = $i; break }
            }
            if ($headerAt -lt 0) { continue }

            $header = $rows[$headerAt]
            $keys   = @{}
            foreach ($k in $header.Keys) { if ($header[$k]) { $keys[$k] = $header[$k] } }

            $records = New-Object System.Collections.Generic.List[object]
            for ($i = $headerAt + 1; $i -lt $rows.Count; $i++) {
                $rec = [ordered]@{}
                $any = $false
                foreach ($k in ($keys.Keys | Sort-Object)) {
                    $v = if ($rows[$i].ContainsKey($k)) { $rows[$i][$k] } else { '' }
                    $rec[$keys[$k]] = $v
                    if ($v) { $any = $true }
                }
                if ($any) { [void]$records.Add($rec) }
            }
            return $records.ToArray()
        }

        throw 'No part number column found in the spreadsheet'
    } finally {
        $zip.Dispose()
        $stream.Dispose()
    }
}


function ConvertFrom-LenovoDelimited {
    <#  CSV or tab separated text to ordered dictionaries, header row detected. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Text)

    $lines = @($Text -split '\r?\n' | Where-Object { $_.Trim() })

    # The header is the first line with several fields that names a part
    # number column; a title line above it is skipped.
    $start = -1
    $delim = ','
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $byTab   = @($lines[$i] -split "`t").Count
        $byComma = @($lines[$i] -split ',').Count
        $d = if ($byTab -gt $byComma) { "`t" } else { ',' }
        if ([Math]::Max($byTab, $byComma) -ge 2 -and $lines[$i] -match $script:PartsHeaderPattern) {
            $start = $i; $delim = $d; break
        }
    }
    if ($start -lt 0) { return @() }

    $objects = @($lines[$start..($lines.Count - 1)] | ConvertFrom-Csv -Delimiter $delim)
    $records = foreach ($o in $objects) {
        $rec = [ordered]@{}
        foreach ($p in $o.PSObject.Properties) { $rec[$p.Name] = [string]$p.Value }
        $rec
    }
    return @($records)
}


function ConvertFrom-LenovoPartsJson {
    <#
        Pulls part records out of whatever JSON shape comes back. Rather than
        guess where the list lives, it walks the whole tree and keeps every
        object that carries a part number: a bare array, an array under
        data/parts, or parts grouped under commodities all come out the same.
        A group's label (commodity, category, tier) is inherited by the parts
        beneath it. A part is an object with a part-number-looking value
        under a part-number-ish key, or - as Lenovo's own list has it - a
        plain id that looks like one next to a name or commodity. Inside a
        part, a nested object or the first object of a nested list is kept
        both as one value (its part number or name) and as its own fields,
        dotted, so a description tucked under localizations.name is still
        found by name; the rest of a nested list collapses to labels.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Text)

    $parsed  = @(ConvertFrom-LenovoJsonList -Text $Text)
    $records = New-Object System.Collections.Generic.List[object]

    $pnKey = $script:PartNumberKeyPattern
    $pnVal = $script:PartNumberValuePattern
    $inheritKey = '(?i)commodity|categor|group|tier|section|class|family'

    $isScalar = { param($v) ($v -is [string]) -or ($v -is [ValueType]) }

    $isPart = {
        param($o)
        $idLooksRight = $false
        $partish      = $false
        foreach ($p in $o.PSObject.Properties) {
            $scalar = ($null -ne $p.Value) -and (& $isScalar $p.Value)
            if ($p.Name -match $pnKey -and $scalar -and ([string]$p.Value).Trim() -match $pnVal) { return $true }
            if ($p.Name -eq 'id') {
                if ($scalar -and ([string]$p.Value).Trim() -match $pnVal) { $idLooksRight = $true }
            } elseif ($p.Name -match '(?i)^name$|desc|commodit|cru|price|quantit|substitut') { $partish = $true }
        }
        return ($idLooksRight -and $partish)
    }

    $labelOf = {
        # One object standing in for a value: its part number, else its name.
        param($o)
        foreach ($p in $o.PSObject.Properties) {
            if ($p.Name -match $pnKey -and $null -ne $p.Value -and (& $isScalar $p.Value) -and [string]$p.Value) { return ([string]$p.Value).Trim() }
        }
        foreach ($p in $o.PSObject.Properties) {
            if ($p.Name -eq 'id' -and $null -ne $p.Value -and (& $isScalar $p.Value) -and ([string]$p.Value).Trim() -match $pnVal) { return ([string]$p.Value).Trim() }
        }
        foreach ($p in $o.PSObject.Properties) {
            if ($p.Name -match '(?i)name|desc|title' -and $p.Value -is [string] -and $p.Value) { return $p.Value.Trim() }
        }
        return ''
    }

    $hoist = {
        # A nested object's own scalar fields, dotted under the parent key.
        param([System.Collections.Specialized.OrderedDictionary] $Rec, [string] $Prefix, $Obj)
        foreach ($q in $Obj.PSObject.Properties) {
            if ($null -ne $q.Value -and (& $isScalar $q.Value)) { $Rec["$Prefix.$($q.Name)"] = [string]$q.Value }
        }
    }

    $flatten = {
        param($o, [hashtable] $Inherited)
        $rec = [ordered]@{}
        foreach ($p in $o.PSObject.Properties) {
            $v = $p.Value
            if ($null -eq $v) { $rec[$p.Name] = ''; continue }
            if (& $isScalar $v) { $rec[$p.Name] = [string]$v; continue }
            if ($v -is [array]) {
                $items = foreach ($i in $v) {
                    if ($null -eq $i) { continue }
                    if (& $isScalar $i) { [string]$i } elseif ($i -is [psobject]) { & $labelOf $i }
                }
                $rec[$p.Name] = (@($items | Where-Object { $_ }) -join ', ')
                if ($p.Name -notmatch '(?i)substitut|\bsubs?\b|bundle|products?$|image') {
                    foreach ($i in $v) {
                        if ($null -ne $i -and -not (& $isScalar $i) -and $i -is [psobject]) { & $hoist $rec $p.Name $i; break }
                    }
                }
                continue
            }
            if ($v -is [psobject]) {
                $rec[$p.Name] = & $labelOf $v
                & $hoist $rec $p.Name $v
            }
        }
        foreach ($k in $Inherited.Keys) {
            if (-not $rec.Contains($k)) { $rec[$k] = $Inherited[$k] }
        }
        return $rec
    }

    $walk = $null
    $walk = {
        param($Node, [hashtable] $Inherited, [int] $Depth)
        if ($null -eq $Node -or $Depth -gt 10) { return }
        if (& $isScalar $Node) { return }

        if ($Node -is [array]) {
            foreach ($item in $Node) { & $walk $item $Inherited ($Depth + 1) }
            return
        }
        if ($Node -is [psobject]) {
            if (& $isPart $Node) {
                [void]$records.Add((& $flatten $Node $Inherited))
                return                       # substitutes inside it are not more parts
            }
            # A grouping object: carry its label down, then look inside.
            $inh = @{} + $Inherited
            foreach ($p in $Node.PSObject.Properties) {
                if ($p.Name -match $inheritKey -and $p.Value -is [string] -and $p.Value) { $inh[$p.Name] = $p.Value.Trim() }
            }
            foreach ($p in $Node.PSObject.Properties) {
                # A substitute or bundle list under something that is not
                # itself a part would come out as parts of this machine.
                if ($p.Name -match '(?i)substitut|bundle') { continue }
                if ($null -ne $p.Value -and -not (& $isScalar $p.Value)) { & $walk $p.Value $inh ($Depth + 1) }
            }
        }
    }

    & $walk $parsed @{} 0
    return $records.ToArray()
}


function Get-LenovoPartColumns {
    <#
        Works out which keys of a set of part records carry which field -
        part number, description, commodity, CRU, price - by wording
        rather than position. Each field gets every candidate key, best
        first, so a record whose first-choice key is empty falls back to the
        next one. The keys seen and a sample value of each come back too.
    #>
    [CmdletBinding()]
    param([object[]] $Record)

    $keys   = New-Object System.Collections.Generic.List[string]
    $sample = @{}
    foreach ($rec in $Record) {
        if ($null -eq $rec) { continue }
        $isDict = $rec -is [System.Collections.IDictionary]
        $names  = if ($isDict) { @($rec.Keys) } else { @($rec.PSObject.Properties | ForEach-Object { $_.Name }) }
        foreach ($n in $names) {
            $k = [string]$n
            if (-not $sample.ContainsKey($k)) { $sample[$k] = ''; [void]$keys.Add($k) }
            if ($sample[$k] -eq '') {
                $v = if ($isDict) { $rec[$n] } else { $rec.$n }
                if ($null -ne $v) { $sample[$k] = ([string]$v).Trim() }
            }
        }
    }

    $pick = {
        # Every key matching any pattern, in pattern order, minus the ones the
        # exclusion rules out (an image URL or a count is never the field).
        param([string[]] $Patterns, [string] $Exclude)
        $found = New-Object System.Collections.Generic.List[string]
        foreach ($pat in $Patterns) {
            foreach ($k in $keys) {
                if ($k -match $pat -and -not ($Exclude -and $k -match $Exclude) -and -not $found.Contains($k)) { [void]$found.Add($k) }
            }
        }
        return $found.ToArray()
    }

    # Keys that belong to something next to the part - a substitute, a
    # bundle, the product - never stand in for the part's own field.
    $other = 'subst|\bsubs?\b|replac|alternat|supersed|parent|bundle|product|image|url|file'
    $map = [ordered]@{
        PartNumber  = @(& $pick @('^\s*fru\s*(part)?\s*(number|no\.?|#|p/?n)?\s*$', 'part[s\s_-]*(number|no\b|no\.|#|num)', '^\s*p/?n\s*$', 'fru', 'partnumber|partno|partnum', '^\s*part\s*$', '^\s*number\s*$', '^\s*id\s*$') `
                                "$other|desc|name|count|total|list|qty|quantity|^(not|is|has)[A-Z_]")
        Description = @(& $pick @('^\s*name\s*$', 'desc', 'part[s\s_-]*name', 'title', '(fru|item|material)[\s_-]*name', 'name', 'label', '^text$|[\s_.-]text$', 'summary') `
                                "$other|commodit|categor|group|class|family|tier|section|model|serial|machine|brand|key|code|id$|html|short|user|operator|host|request")
        Commodity   = @(& $pick @('commodity[\s_.-]*(name|desc|label|display|text|val)', 'commodity', 'categor', 'part[s\s_-]*type', '^\s*type\s*$', 'group', 'class', 'family') `
                                "$other|id$|count|total|filter|list|param|key|code|lang|source")
        Cru         = @(& $pick @('^\s*cru', 'cru[\s_-]*(tier|type|level)', '\bcru\b', 'serviceab') $other)
        Price       = @(& $pick @('price', 'cost') "$other|currency|symbol|unit|type|id$|format")
    }

    return [pscustomobject]@{ Keys = $keys.ToArray(); Sample = $sample; Map = $map }
}


function ConvertTo-LenovoCruName {
    <#
        Lenovo's CRU tier codes, worded the way its parts page words them:
        who may replace the part. Anything that is not one of the codes
        passes through untouched.
    #>
    [CmdletBinding()]
    param([string] $Code)

    switch (([string]$Code).Trim()) {
        '0'     { return 'Consumable' }
        '1'     { return 'CRU T1 (mandatory)' }
        '2'     { return 'CRU T2 (optional)' }
        '9'     { return 'FRU only' }
        default { return ([string]$Code).Trim() }
    }
}


function ConvertTo-LenovoPartRows {
    <#
        Maps whatever column names the export uses onto the fixed set the
        tool reports: PartNumber, Description, Commodity, Cru, Price.
        Column matching is by wording (Get-LenovoPartColumns) so a renamed
        header or a reordered sheet does not break anything, and each field
        takes the first of its candidate keys that holds a value in that
        record. A record with no part number at all is dropped: the part
        number is the point.
    #>
    [CmdletBinding()]
    param([object[]] $Record)

    if ($null -eq $Record -or $Record.Count -eq 0) { return @() }

    $cols = Get-LenovoPartColumns -Record $Record
    $map  = $cols.Map
    if (@($map['PartNumber']).Count -eq 0) { throw "No part number column in: $($cols.Keys -join ', ')" }

    $get = {
        param($Rec, [string] $Key)
        if (-not $Key) { return '' }
        $v = if ($Rec -is [System.Collections.IDictionary]) { $Rec[$Key] }
             elseif ($Rec.PSObject.Properties[$Key]) { $Rec.$Key } else { $null }
        if ($null -eq $v) { return '' }
        return ([string]$v).Trim()
    }
    $first = {
        param($Rec, [string[]] $Keys)
        foreach ($k in $Keys) { $v = & $get $Rec $k; if ($v) { return $v } }
        return ''
    }

    $rows = foreach ($rec in $Record) {
        $pn = (& $first $rec $map['PartNumber']) -replace '\s+', ''
        if (-not $pn) {
            # Nothing under the usual names: take any part-number-looking
            # value under a part-number-ish key, which is what got the
            # record kept in the first place.
            $names = if ($rec -is [System.Collections.IDictionary]) { @($rec.Keys) } else { @($rec.PSObject.Properties | ForEach-Object { $_.Name }) }
            foreach ($k in $names) {
                if ([string]$k -match $script:PartNumberKeyPattern) {
                    $v = (& $get $rec ([string]$k)) -replace '\s+', ''
                    if ($v -match $script:PartNumberValuePattern) { $pn = $v; break }
                }
            }
        }
        if (-not $pn) { continue }
        [pscustomobject]@{
            PartNumber  = $pn.ToUpperInvariant()
            Description = & $first $rec $map['Description']
            Commodity   = & $first $rec $map['Commodity']
            Cru         = ConvertTo-LenovoCruName -Code (& $first $rec $map['Cru'])
            Price       = & $first $rec $map['Price']
        }
    }
    return @($rows)
}


function Get-LenovoJsonShape {
    <#
        A compact sketch of a JSON document - keys, nesting, array sizes and
        a sample of each value - for the attempt log. Enough to see where a
        list lives and what its fields are called without dumping it all.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Text, [int] $MaxLength = 900)

    $sketch = $null
    $sketch = {
        param($Node, [int] $Depth)
        if ($null -eq $Node) { return 'null' }
        if ($Node -is [string]) {
            $t = $Node
            if ($t.Length -gt 36) { $t = $t.Substring(0, 36) + '...' }
            return '"' + $t + '"'
        }
        if ($Node -is [ValueType]) { return [string]$Node }
        if ($Depth -gt 6) { return '...' }
        if ($Node -is [array]) {
            if ($Node.Count -eq 0) { return '[]' }
            return "[$($Node.Count): " + (& $sketch $Node[0] ($Depth + 1)) + ']'
        }
        if ($Node -is [psobject]) {
            $parts = @()
            $i = 0
            foreach ($p in $Node.PSObject.Properties) {
                if ($i -ge 60) { $parts += '...'; break }
                $parts += ($p.Name + ':' + (& $sketch $p.Value ($Depth + 1)))
                $i++
            }
            return '{' + ($parts -join ', ') + '}'
        }
        return [string]$Node
    }

    try {
        $parsed = @(ConvertFrom-LenovoJsonList -Text $Text)
        $out = if ($parsed.Count -eq 1) { & $sketch $parsed[0] 0 } else { & $sketch $parsed 0 }
    } catch {
        $out = 'unreadable JSON'
    }
    if ($out.Length -gt $MaxLength) { $out = $out.Substring(0, $MaxLength) + '...' }
    return $out
}


function Get-LenovoCruTiers {
    <#
        Reads the CRU tier codes out of whatever parts/config answered: the
        first list found under a key mentioning "cru", taking each entry's
        value/code/id (or the entry itself when it is a bare string).
    #>
    [CmdletBinding()]
    param([string] $ConfigText)

    if (-not $ConfigText) { return @() }
    try { $parsed = @(ConvertFrom-LenovoJsonList -Text $ConfigText) } catch { return @() }

    $hits = New-Object System.Collections.Generic.List[object]
    $find = $null
    $find = {
        param($Node, [int] $Depth)
        if ($null -eq $Node -or $Depth -gt 8 -or $hits.Count -gt 0) { return }
        if ($Node -is [string] -or $Node -is [ValueType]) { return }
        if ($Node -is [array]) { foreach ($i in $Node) { & $find $i ($Depth + 1) }; return }
        if ($Node -is [psobject]) {
            foreach ($p in $Node.PSObject.Properties) {
                if ($p.Name -match '(?i)cru' -and $p.Value -is [array] -and $p.Value.Count -gt 0) { [void]$hits.Add($p.Value); return }
            }
            foreach ($p in $Node.PSObject.Properties) { & $find $p.Value ($Depth + 1) }
        }
    }
    & $find $parsed 0
    if ($hits.Count -eq 0) { return @() }
    $found = $hits[0]

    $tiers = foreach ($entry in $found) {
        if ($null -eq $entry) { continue }
        if ($entry -is [string] -or $entry -is [ValueType]) { [string]$entry; continue }
        $v = $null
        foreach ($k in 'value', 'code', 'id', 'key', 'tier', 'cruTier', 'name') {
            if ($entry.PSObject.Properties[$k] -and $null -ne $entry.$k -and [string]$entry.$k) { $v = [string]$entry.$k; break }
        }
        if ($v) { $v }
    }
    return @($tiers | Where-Object { $_ } | Select-Object -Unique)
}


function Read-LenovoPartsPayload {
    <#
        Works out what kind of thing came back - spreadsheet, JSON, CSV, a web
        page, or a one line brush-off - and reads the part records out of it.
        Reason is set, and Records left empty, when it is not a parts list.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Response)

    $out = [ordered]@{ Source = ''; Records = @(); Reason = '' }
    $b = $Response.Bytes
    if ($null -eq $b -or $b.Length -eq 0) {
        $out.Reason = 'empty reply'
        return [pscustomobject]$out
    }

    try {
        if ($b.Length -ge 2 -and $b[0] -eq 0x50 -and $b[1] -eq 0x4B) {
            $out.Source  = 'xlsx'
            $out.Records = @(ConvertFrom-LenovoXlsx -Bytes $b)
        } else {
            $text = $Response.Text.TrimStart([char]0xFEFF, ' ', "`t", "`r", "`n")
            if ($text -match '^[\[{]') {
                $out.Source  = 'json'
                $out.Records = @(ConvertFrom-LenovoPartsJson -Text $text)
                if ($out.Records.Count -eq 0) {
                    # A well-formed refusal: {"code":103,"msg":{"desc":"..."}}
                    $top = @(ConvertFrom-LenovoJsonList -Text $text)
                    if ($top.Count -eq 1 -and $top[0] -is [psobject] -and $top[0].PSObject.Properties['code'] -and
                        [string]$top[0].code -ne '0') {
                        $desc = ''
                        if ($top[0].PSObject.Properties['msg'] -and $top[0].msg) {
                            $desc = if ($top[0].msg -is [string]) { $top[0].msg }
                                    elseif ($top[0].msg.PSObject.Properties['desc']) { [string]$top[0].msg.desc } else { '' }
                        }
                        if (-not $desc -and $top[0].PSObject.Properties['message']) { $desc = [string]$top[0].message }
                        if (-not $desc) { $desc = "code $($top[0].code)" }
                        $out.Reason = "Lenovo said: $($desc.Trim())"
                        return [pscustomobject]$out
                    }
                }
            } elseif ($text -match '^<') {
                $out.Reason = 'a web page, not a parts list'
                return [pscustomobject]$out
            } elseif ($text -notmatch $script:PartsHeaderPattern) {
                $short = ($text -replace '\s+', ' ').Trim()
                if ($short.Length -gt 80) { $short = $short.Substring(0, 80) + '...' }
                $out.Reason = "not a parts list: $short"
                return [pscustomobject]$out
            } else {
                $out.Source  = 'csv'
                $out.Records = @(ConvertFrom-LenovoDelimited -Text $text)
            }
        }
    } catch {
        $out.Reason = "could not read it: $($_.Exception.Message)"
        return [pscustomobject]$out
    }

    if ($out.Records.Count -eq 0) { $out.Reason = "no part rows in the $($out.Source)" }
    return [pscustomobject]$out
}


function Get-LenovoPartsList {
    <#
    .SYNOPSIS
        Downloads the full FRU parts list for one Lenovo serial number.

    .DESCRIPTION
        Resolves the serial to its machine type and model, then pulls the
        parts list Lenovo's parts lookup page offers as "Download parts list".
        Returns one object carrying the product details and a Parts array of
        PartNumber / Description / Commodity / Cru / Price rows.
        Problems are reported in Error rather than thrown.

    .EXAMPLE
        (Get-LenovoPartsList PF0ABCDE).Parts | Format-Table
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)] [string] $SerialNumber,
        [int] $TimeoutSec = 60
    )

    $serial = $SerialNumber.Trim().ToUpperInvariant()
    $out = [ordered]@{
        Serial      = $serial
        Product     = ''
        MachineType = ''
        Model       = ''
        ProductId   = ''
        Parts       = @()
        Source      = ''
        Error       = ''
        Raw         = ''      # what Lenovo sent, kept when something goes wrong
    }

    $client = New-LenovoHttpClient -TimeoutSec $TimeoutSec
    try {
        $product = Get-LenovoProduct -SerialNumber $serial -Client $client
        $out.Product     = $product.Product
        $out.MachineType = $product.MachineType
        $out.Model       = $product.Model
        $out.ProductId   = $product.ProductId
        if ($product.Error) {
            $out.Error = $product.Error
            $out.Raw   = $product.Raw
            return [pscustomobject]$out
        }

        $stamp    = Get-Date -Format 'yyyy-MM-dd-HH-mm-ss'
        $fileName = "PartsExport_Serial-$($serial.ToLowerInvariant())_$stamp.xlsx"
        $fields   = [ordered]@{
            type         = 'SERIAL'
            serialId     = $serial.ToLowerInvariant()
            model        = $product.Model.ToLowerInvariant()
            mtId         = $product.MachineType.ToLowerInvariant()
            supportSales = 'true'
            viewInStock  = 'false'
        }
        $query = '?fileName=' + [Uri]::EscapeDataString($fileName) + '&' +
                 (($fields.Keys | ForEach-Object { "$_=" + [Uri]::EscapeDataString($fields[$_]) }) -join '&')
        $accept   = 'application/json, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, text/csv, */*'
        $pageUrl  = "$($script:SiteBase)/us/en/products/$($product.ProductId)/parts"
        $xhr      = @{ 'X-Requested-With' = 'XMLHttpRequest'; 'Origin' = $script:SiteBase }

        # The body the parts page itself posts, read out of its scripts. The
        # page also sends the CRU tiers it is filtering on; those come from
        # parts/config when it lists them, else the usual codings are tried.
        $base = [ordered]@{
            serialId       = $fields.serialId
            mtId           = $fields.mtId
            model          = $fields.model
            couponNumber   = ''
            source         = ''
            channel        = ''
            firstGenCPU    = $false
            fetchProcessor = $true
        }
        $withTiers = {
            param([string[]] $Tiers)
            $b = [ordered]@{}
            foreach ($k in $base.Keys) { $b[$k] = $base[$k] }
            if ($Tiers -and $Tiers.Count -gt 0) { $b['cruTiers'] = @($Tiers) }
            return (ConvertTo-Json -InputObject $b -Compress -Depth 5)
        }

        $log    = New-Object System.Collections.Generic.List[string]
        $config = Invoke-LenovoRequest -Client $client -Url "$($script:PartsApiBase)/config" -Method POST `
                                       -Accept $accept -Referer $pageUrl -Headers $xhr -Body '{}'
        $tiers  = @()
        if (-not $config.Error -and $config.Text) { $tiers = @(Get-LenovoCruTiers -ConfigText $config.Text) }
        [void]$log.Add("0. POST parts/config -> HTTP $($config.StatusCode) $($config.ContentType); CRU tiers found: " +
                       $(if ($tiers.Count -gt 0) { $tiers -join ',' } else { 'none' }) +
                       $(if ($config.Text -and $config.Text -match '^\s*[\[{]') { '; shape: ' + (Get-LenovoJsonShape -Text $config.Text -MaxLength 500) } else { '' }))

        $tierSets = New-Object System.Collections.Generic.List[object]
        [void]$tierSets.Add(@())                                           # exactly what was observed first
        if ($tiers.Count -gt 0) { [void]$tierSets.Add($tiers) }            # what the site's config lists
        foreach ($guess in @(@('1', '2', '3'), @('S', 'O', 'N'))) {        # the usual codings
            if (-not ($tierSets | Where-Object { ($_ -join ',') -eq ($guess -join ',') })) { [void]$tierSets.Add($guess) }
        }
        $bestTiers = if ($tiers.Count -gt 0) { $tiers } else { @('1', '2', '3') }

        # As-built is the list for this very serial, so every tier coding is
        # tried there first; model and compatible are broader fallbacks, and
        # the page's "Download parts list" export is the last resort. What
        # each one answered is kept in Raw.
        $attempts = @()
        foreach ($set in $tierSets) {
            $label = if ($set.Count -gt 0) { " with cruTiers $($set -join ',')" } else { '' }
            $attempts += @{ Name = "POST parts/asBuilt$label"; Method = 'POST'; Url = "$($script:PartsApiBase)/asBuilt"; Body = (& $withTiers $set) }
        }
        $attempts += @{ Name = "POST parts/model with cruTiers $($bestTiers -join ',')";      Method = 'POST'; Url = "$($script:PartsApiBase)/model";      Body = (& $withTiers $bestTiers) }
        $attempts += @{ Name = "POST parts/compatible with cruTiers $($bestTiers -join ',')"; Method = 'POST'; Url = "$($script:PartsApiBase)/compatible"; Body = (& $withTiers $bestTiers) }
        $attempts += @{ Name = 'POST parts/export with query string';                        Method = 'POST'; Url = $script:PartsExportUrl + $query + '&cruTiers=' + [Uri]::EscapeDataString($bestTiers -join ','); Body = '' }

        $payload = $null
        $n = 0
        foreach ($a in $attempts) {
            $n++
            # Indexer access: a missing optional key reads as $null, where dot
            # notation would throw under strict mode.
            $resp = Invoke-LenovoRequest -Client $client -Url $a['Url'] -Method $a['Method'] -Accept $accept `
                                         -Referer $pageUrl -Headers $xhr -Body ([string]$a['Body'])
            if ($resp.Error -and $resp.StatusCode -eq 0) {
                [void]$log.Add("$n. $($a['Name']) -> $($resp.Error)")
                break                                   # no network; nothing else will work either
            }
            if ($resp.Error) {
                [void]$log.Add("$n. $($a['Name']) -> " + (Get-LenovoResponseSummary $resp))
                continue
            }
            $read = Read-LenovoPartsPayload -Response $resp
            if ($read.Records.Count -gt 0) {
                [void]$log.Add("$n. $($a['Name']) -> HTTP $($resp.StatusCode) $($read.Source), $($read.Records.Count) records")
                $payload = $read
                break
            }
            $detail = ''
            if ($read.Source -eq 'json' -and $read.Reason -notmatch '^Lenovo said') {
                # What came back, sketched, so an unexpected layout shows in
                # the error rather than a dead end.
                $detail = '; shape: ' + (Get-LenovoJsonShape -Text $resp.Text)
            }
            [void]$log.Add("$n. $($a['Name']) -> HTTP $($resp.StatusCode) $($resp.ContentType): $($read.Reason)$detail")
        }
        $out.Raw = ($log -join "`n")

        if (-not $payload) {
            # The first real attempt is the site's own request shape, so its
            # answer is the one worth repeating; the rest are in Raw.
            $first = [string]$log[1]
            if ($first -match 'not a parts list: (.+)$')        { $out.Error = "Lenovo refused the parts list: $($Matches[1])" }
            elseif ($first -match 'Lenovo said: (.+)$')         { $out.Error = "Lenovo said: $($Matches[1])" }
            elseif ($first -match 'no part rows in the json')   { $out.Error = 'Lenovo returned no parts for this serial' }
            elseif ($first -match 'a web page')                 { $out.Error = 'Lenovo returned a web page instead of a parts list' }
            elseif ($first -match '-> (Request failed.*)$')     { $out.Error = $Matches[1] }
            elseif ($first -match '-> HTTP (\d+)')              { $out.Error = "Lenovo did not return a parts list (HTTP $($Matches[1]))" }
            else                                                { $out.Error = 'Lenovo did not return a parts list' }
            return [pscustomobject]$out
        }

        $out.Source = $payload.Source
        $out.Parts  = @(ConvertTo-LenovoPartRows -Record $payload.Records)
        if ($out.Parts.Count -eq 0) { $out.Error = 'No parts listed for this serial' }
    } catch {
        $out.Error = "Could not read the parts list: $($_.Exception.Message)"
    } finally {
        $client.Dispose()
    }

    return [pscustomobject]$out
}


function Find-LenovoPart {
    <#
    .SYNOPSIS
        The part numbers under one of Lenovo's commodities for one serial.

    .DESCRIPTION
        Looks up the serial's parts list and keeps the rows under the chosen
        commodity - Lenovo's own grouping, which differs per machine (see
        Get-LenovoPartCommodity). Returns the parts list object with
        Commodity (as resolved), Commodities (the machine's list), Matches
        and Note added; Note says so when the commodity asked for is not one
        this machine has, and Matches is then empty.

    .EXAMPLE
        (Find-LenovoPart PF0ABCDE 'System boards').Matches
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)] [string] $SerialNumber,
        [Parameter(Position = 1)] [Alias('Part')] [string] $Commodity = 'All parts',
        [int] $TimeoutSec = 60
    )

    $list = Get-LenovoPartsList -SerialNumber $SerialNumber -TimeoutSec $TimeoutSec
    $hits = @()
    $names = @()
    $name  = $Commodity
    $note  = ''
    if (-not $list.Error) {
        $names = @(Get-LenovoPartCommodity -Part $list.Parts)
        if (-not $Commodity -or $Commodity -match '^\s*all(\s*parts)?\s*$') {
            $name = 'All parts'
            $hits = @($list.Parts)
        } else {
            try {
                $name = Resolve-LenovoPartCommodity -Part $list.Parts -Commodity $Commodity
                $hits = @(Select-LenovoPart -Part $list.Parts -Commodity $name)
            } catch {
                $note = $_.Exception.Message
            }
        }
    }

    $list | Add-Member -NotePropertyName Commodity   -NotePropertyValue $name
    $list | Add-Member -NotePropertyName Commodities -NotePropertyValue $names
    $list | Add-Member -NotePropertyName Matches     -NotePropertyValue $hits
    $list | Add-Member -NotePropertyName Note        -NotePropertyValue $note
    return $list
}


Export-ModuleMember -Function Get-LenovoWarranty, ConvertTo-SerialList,
                              Format-WarrantyDate, ConvertFrom-LenovoIbaseInfo,
                              ConvertTo-DateOrNull,
                              Get-LenovoProduct, Get-LenovoPartsList, Find-LenovoPart,
                              Get-LenovoPartCommodity, Resolve-LenovoPartCommodity, Select-LenovoPart,
                              ConvertTo-LenovoPartRows, ConvertFrom-LenovoXlsx,
                              ConvertFrom-LenovoDelimited, ConvertFrom-LenovoPartsJson,
                              ConvertFrom-LenovoJsonList, Read-LenovoPartsPayload,
                              Get-LenovoJsonShape, Get-LenovoCruTiers,
                              Get-LenovoPartColumns, ConvertTo-LenovoCruName
