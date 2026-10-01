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
#    2. /us/en/api/v4/upsellAggregation/parts/export
#         ?type=SERIAL&serialId=pf0abcde&model=21ah00bbus&mtId=21ah
#       The "Download parts list" link on Lenovo's own parts lookup page.
#       It hands back the full FRU list for that serial as a spreadsheet.
#
#  The spreadsheet is read straight out of the xlsx (it is just a zip of
#  XML) so there is no Excel dependency. The normaliser also accepts CSV
#  or JSON, so if Lenovo changes what the export returns the rest of the
#  tool keeps working as long as there is still a part number column.
# ==========================================================================

$script:SiteBase       = 'https://pcsupport.lenovo.com'
$script:ProductsUrl    = 'https://pcsupport.lenovo.com/us/en/api/v4/mse/getproducts'
$script:PartsExportUrl = 'https://pcsupport.lenovo.com/us/en/api/v4/upsellAggregation/parts/export'
$script:PartsReferer   = 'https://pcsupport.lenovo.com/us/en/partslookup'

# What a parts list header row looks like, used to skip any title lines
# Lenovo puts above it.
$script:PartsHeaderPattern = 'part\s*(number|no\b|no\.|#|num)|\bfru\b|descr'

Add-Type -AssemblyName System.IO.Compression            -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue


# The preset list the user picks from. Include is matched against the part
# description and commodity; Exclude is matched against the description only
# and knocks out the brackets, cables and screws that mention the same word.
$script:PartCategories = @(
    @{ Name = 'LCD panel';          Include = 'LCD|DISPLAY|PANEL|SCREEN';
                                    Exclude = 'COVER|BEZEL|CABLE|HINGE|BRACKET|TAPE|FOIL|RUBBER|SCREW|FRAME|SUPPORT|SPONGE|GASKET|PROTECT|FILM|FELT|MYLAR|SHIELD' }
    @{ Name = 'LCD back cover';     Include = 'REAR\s*COVER|BACK\s*COVER|LCD\s*COVER|COVER[\s,_]*LCD\s*REAR|\bA[\s_-]?COVER\b|REAR\s*CASE|LCD\s*BACK|LCD\s*REAR';
                                    Exclude = 'BEZEL|HINGE\s*(CAP|COVER)|SCREW|RUBBER|TAPE|FOOT' }
    @{ Name = 'LCD bezel';          Include = 'BEZEL|\bB[\s_-]?COVER\b|LCD\s*FRONT|FRONT\s*COVER';
                                    Exclude = 'KEYBOARD|\bKBD\b|PALM|SCREW|TAPE' }
    @{ Name = 'LCD cable';          Include = 'LCD\s*CABLE|CABLE[\s,_]*LCD|\bEDP\b|LVDS|DISPLAY\s*CABLE|PANEL\s*CABLE';
                                    Exclude = 'SCREW|TAPE' }
    @{ Name = 'Hinges';             Include = 'HINGE';
                                    Exclude = 'CAP|SCREW|RUBBER|TAPE' }
    @{ Name = 'System board';       Include = 'SYSTEM\s*BOARD|PLANAR|MAIN\s*BOARD|MOTHER\s*BOARD|\bMB\b';
                                    Exclude = 'CABLE|BRACKET|SHIELD|SCREW|STANDOFF|INSULAT|MYLAR|TAPE|FOAM' }
    @{ Name = 'Power button board'; Include = 'POWER\s*(BUTTON|SWITCH|BOARD)|PWR\s*(BTN|BUTTON|BOARD)';
                                    Exclude = 'CABLE|FFC|SCREW|SUPPLY|ADAPTER' }
    @{ Name = 'I/O board';          Include = 'I/?O\s*BOARD|USB\s*BOARD|SUB\s*BOARD|SUBCARD|DAUGHTER|AUDIO\s*BOARD';
                                    Exclude = 'CABLE|FFC|SCREW|BRACKET' }
    @{ Name = 'SSD';                Include = '\bSSD\b|SOLID[\s-]*STATE|NVME|M\.2|PCIE.*(DRIVE|STORAGE)';
                                    Exclude = 'BRACKET|THERMAL|CABLE|TRAY|SCREW|PAD|RUBBER|SHIELD|CADDY|FOAM|HOLDER|DOOR|COVER' }
    @{ Name = 'Hard drive';         Include = '\bHDD\b|HARD\s*(DISK|DRIVE)|SATA.*(DRIVE|HDD)';
                                    Exclude = 'BRACKET|CABLE|TRAY|SCREW|RUBBER|RAIL|CADDY|FOAM|HOLDER|DOOR|COVER' }
    @{ Name = 'Memory';             Include = 'MEMORY|\bDIMM\b|SODIMM|\bRAM\b|DDR\d';
                                    Exclude = 'COVER|DOOR|SHIELD|BRACKET|SCREW|FOAM|MYLAR' }
    @{ Name = 'Battery';            Include = 'BATTERY|\bBATT\b|\bBTY\b';
                                    Exclude = 'CABLE|COVER|BRACKET|SCREW|TAPE|CMOS|\bRTC\b|COIN|BACKUP|FOAM' }
    @{ Name = 'AC adapter';         Include = 'ADAPTER|\bADPT\b|CHARGER|POWER\s*SUPPLY|\bPSU\b';
                                    Exclude = 'CORD|ETHERNET|HDMI|VGA|DONGLE|DISPLAYPORT|\bDP\b|USB-?C\s*TO|BRACKET' }
    @{ Name = 'Power cord';         Include = 'POWER\s*CORD|LINE\s*CORD|\bCORD\b|AC\s*CABLE';
                                    Exclude = 'BRACKET' }
    @{ Name = 'Keyboard';           Include = 'KEYBOARD|\bKBD\b|\bKB\b|\bKYB\b';
                                    Exclude = 'BEZEL|COVER|PALM|BRACKET|PLATE|CABLE|PROTECT|SKIN|SHIELD|TAPE|FOAM|SCREW' }
    @{ Name = 'Palmrest / C cover'; Include = 'PALM\s*REST|PALMREST|KEYBOARD\s*BEZEL|KBD\s*BEZEL|\bC[\s_-]?COVER\b|UPPER\s*CASE|TOP\s*CASE|KBD\s*COVER';
                                    Exclude = 'SCREW|TAPE|FOAM|MYLAR' }
    @{ Name = 'Base cover / D cover'; Include = 'BASE\s*COVER|BOTTOM\s*(COVER|CASE)|\bD[\s_-]?COVER\b|LOWER\s*CASE|BASE\s*(ASM|ASSY|ASSEMBLY|ENCLOSURE)';
                                    Exclude = 'FOOT|FEET|RUBBER|SCREW|DOOR|TAPE' }
    @{ Name = 'Touchpad';           Include = 'TOUCH\s*PAD|TRACK\s*PAD|CLICK\s*PAD';
                                    Exclude = 'CABLE|BRACKET|\bFFC\b|MYLAR|SCREW|TAPE' }
    @{ Name = 'Fingerprint reader'; Include = 'FINGER\s*PRINT|\bFPR\b|\bFP\s*(READER|SENSOR|BOARD|MODULE)';
                                    Exclude = 'CABLE|BRACKET|\bFFC\b|SCREW' }
    @{ Name = 'Fan / heatsink';     Include = '\bFAN\b|HEAT\s*SINK|HEATSINK|THERMAL|COOLER|COOLING';
                                    Exclude = 'PAD|PASTE|GREASE|TAPE|SHEET|FOAM|GRAPHITE|BRACKET|CABLE|SCREW' }
    @{ Name = 'Wireless card';      Include = '\bWLAN\b|\bWWAN\b|WIRELESS|WI-?FI|BLUETOOTH|\bLTE\b|\b5G\b|INTEL\s*AX\d';
                                    Exclude = 'ANTENNA|CABLE|BRACKET|COVER|TRAY|\bSIM\b|SCREW|TAPE' }
    @{ Name = 'Antenna';            Include = 'ANTENNA';
                                    Exclude = 'SCREW|TAPE' }
    @{ Name = 'Camera';             Include = 'CAMERA|WEBCAM|\bCAM\b|\bIR\s*CAM';
                                    Exclude = 'SHUTTER|CABLE|BRACKET|TAPE|FOAM|MYLAR|SCREW' }
    @{ Name = 'Speakers';           Include = 'SPEAKER|\bSPK\b';
                                    Exclude = 'CABLE|BRACKET|GRILL|FOAM|TAPE|SCREW' }
    @{ Name = 'Screws';             Include = 'SCREW';
                                    Exclude = '' }
) | ForEach-Object { [pscustomobject]$_ }


function Get-LenovoPartCategory {
    <#
    .SYNOPSIS
        The preset list of parts the tool knows how to pick out of a parts list.
    .DESCRIPTION
        Returns one object per preset with Name, Include and Exclude (the
        regular expressions used to match it). Pass -Name to get just one.
    #>
    [CmdletBinding()]
    param([string] $Name)

    if (-not $Name) { return $script:PartCategories }

    $hit = $script:PartCategories | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
    if ($hit) { return $hit }

    # Be forgiving about punctuation and spacing: "lcd back cover", "Base cover".
    $flat = ($Name -replace '[^A-Za-z0-9]', '').ToUpperInvariant()
    $hit = $script:PartCategories | Where-Object {
        (($_.Name -replace '[^A-Za-z0-9]', '').ToUpperInvariant()) -eq $flat
    } | Select-Object -First 1
    if ($hit) { return $hit }

    $names = ($script:PartCategories | ForEach-Object { $_.Name }) -join ', '
    throw "Unknown part '$Name'. Choose one of: $names"
}


function Select-LenovoPart {
    <#
    .SYNOPSIS
        Filters a parts list down to one preset part.
    .DESCRIPTION
        Matching is by wording: the preset's Include pattern is tested against
        the description and commodity, then its Exclude pattern knocks out
        brackets, cables and screws that happen to mention the same word. An
        empty Category or 'All parts' returns the list untouched.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)] [object[]] $Part,
        [Parameter(Position = 1)] [string]   $Category
    )

    if ($null -eq $Part) { return @() }
    if (-not $Category -or $Category -match '^\s*all(\s*parts)?\s*$') { return @($Part) }

    $cat = Get-LenovoPartCategory -Name $Category

    return @($Part | Where-Object {
        $desc = [string]$_.Description
        $both = $desc + ' | ' + [string]$_.Commodity
        ($both -match $cat.Include) -and
        (-not $cat.Exclude -or ($desc -notmatch $cat.Exclude))
    })
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
    <#  One line saying what Lenovo actually sent, for the diagnostics. #>
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
        Finds the list of parts inside whatever JSON shape comes back - a bare
        array, or an array under data/parts/items/results a few levels down.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $Text)

    $parsed = @(ConvertFrom-LenovoJsonList -Text $Text)

    $find = $null
    $find = {
        param($Node, [int] $Depth)
        if ($null -eq $Node -or $Depth -gt 5) { return $null }

        if ($Node -is [array]) {
            $objs = @($Node | Where-Object { $_ -is [psobject] -and -not ($_ -is [string]) })
            if ($objs.Count -gt 0) {
                $names = @($objs[0].PSObject.Properties | ForEach-Object { $_.Name })
                if (($names -join ' ') -match 'part|fru|\bp/?n\b') { return ,$objs }
            }
            foreach ($item in $Node) {
                $hit = & $find $item ($Depth + 1)
                if ($hit) { return $hit }
            }
            return $null
        }

        if ($Node -is [psobject]) {
            # Likely wrappers first, then everything else.
            $props = @($Node.PSObject.Properties)
            $ordered = @($props | Where-Object { $_.Name -match '^(parts?|data|items?|results?|list|rows|records|content|body)$' }) +
                       @($props | Where-Object { $_.Name -notmatch '^(parts?|data|items?|results?|list|rows|records|content|body)$' })
            foreach ($p in $ordered) {
                $hit = & $find $p.Value ($Depth + 1)
                if ($hit) { return $hit }
            }
        }
        return $null
    }

    $list = & $find $parsed 0
    if (-not $list) { return @() }

    $records = foreach ($o in $list) {
        $rec = [ordered]@{}
        foreach ($p in $o.PSObject.Properties) {
            $v = $p.Value
            if ($v -is [array]) { $v = ($v | ForEach-Object { [string]$_ }) -join ', ' }
            $rec[$p.Name] = [string]$v
        }
        $rec
    }
    return @($records)
}


function ConvertTo-LenovoPartRows {
    <#
        Maps whatever column names the export uses onto the fixed set the
        tool reports: PartNumber, Description, Commodity, Substitutes, Status,
        Cru, Price. Column matching is by wording so a renamed header or a
        reordered sheet does not break anything.
    #>
    [CmdletBinding()]
    param([object[]] $Record)

    if ($null -eq $Record -or $Record.Count -eq 0) { return @() }

    $first = $Record[0]
    $keys  = @(if ($first -is [System.Collections.IDictionary]) { $first.Keys } else { $first.PSObject.Properties | ForEach-Object { $_.Name } })

    $pick = {
        param([string[]] $Patterns)
        foreach ($pat in $Patterns) {
            foreach ($k in $keys) {
                if (([string]$k) -match $pat) { return [string]$k }
            }
        }
        return $null
    }

    $map = @{
        PartNumber  = & $pick @('^\s*fru\s*(part)?\s*(number|no\.?|#|p/?n)?\s*$', 'part\s*(number|no\b|no\.|#|num)', '^\s*p/?n\s*$', 'fru', 'partnumber|partno|partnum', '^\s*part\s*$', '^\s*number\s*$')
        Description = & $pick @('desc', 'part\s*name', '^\s*name\s*$', 'title')
        Commodity   = & $pick @('commodity', 'categor', 'part\s*type', '^\s*type\s*$', 'group', 'class')
        Substitutes = & $pick @('subst', 'replac', 'alternat', 'supersed')
        Status      = & $pick @('status', 'availab', 'stock', 'orderable')
        Cru         = & $pick @('\bcru\b')
        Price       = & $pick @('price', 'cost')
    }
    if (-not $map.PartNumber) { throw "No part number column in: $($keys -join ', ')" }

    $get = {
        param($Rec, [string] $Key)
        if (-not $Key) { return '' }
        $v = if ($Rec -is [System.Collections.IDictionary]) { $Rec[$Key] } else { $Rec.$Key }
        if ($null -eq $v) { return '' }
        return ([string]$v).Trim()
    }

    $rows = foreach ($rec in $Record) {
        $pn = (& $get $rec $map.PartNumber) -replace '\s+', ''
        $ds = & $get $rec $map.Description
        if (-not $pn -and -not $ds) { continue }
        [pscustomobject]@{
            PartNumber  = $pn.ToUpperInvariant()
            Description = $ds
            Commodity   = & $get $rec $map.Commodity
            Substitutes = & $get $rec $map.Substitutes
            Status      = & $get $rec $map.Status
            Cru         = & $get $rec $map.Cru
            Price       = & $get $rec $map.Price
        }
    }
    return @($rows)
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
        PartNumber / Description / Commodity / Substitutes / Status rows.
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
        $jsonBody = '{"fileName":"' + $fileName + '","type":"SERIAL","serialId":"' + $fields.serialId +
                    '","model":"' + $fields.model + '","mtId":"' + $fields.mtId +
                    '","supportSales":true,"viewInStock":false}'

        $accept   = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/json, text/csv, */*'
        $pageUrl  = "$($script:SiteBase)/us/en/products/$($product.ProductId)/parts"
        $xhr      = @{ 'X-Requested-With' = 'XMLHttpRequest'; 'Origin' = $script:SiteBase }

        # The export is the site's own "Download parts list" link, captured
        # from a browser rather than documented, so the exact shape of the
        # request it wants is not certain. These are tried in turn until one
        # hands back a parts list; what each one answered is kept in Raw.
        $attempts = @(
            @{ Name = 'POST export with query string';      Method = 'POST'; Url = $script:PartsExportUrl + $query; Body = '' }
            @{ Name = 'GET export with query string';       Method = 'GET';  Url = $script:PartsExportUrl + $query }
            @{ Name = 'POST export with JSON body';         Method = 'POST'; Url = $script:PartsExportUrl;          Body = $jsonBody }
            @{ Name = 'POST export after loading the page'; Method = 'POST'; Url = $script:PartsExportUrl + $query; Body = ''; WarmUp = $pageUrl }
        )

        $log     = New-Object System.Collections.Generic.List[string]
        $payload = $null
        $n = 0
        foreach ($a in $attempts) {
            $n++
            # Indexer access: a missing optional key reads as $null, where dot
            # notation would throw under strict mode.
            if ($a['WarmUp']) {
                # Cookies the page sets ride along in the client for the retry.
                $warm = Invoke-LenovoRequest -Client $client -Url $a['WarmUp'] -Accept 'text/html, */*'
                [void]$log.Add("$n. GET parts page -> HTTP $($warm.StatusCode) $($warm.ContentType) $($warm.Bytes.Length) bytes")
            }
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
            [void]$log.Add("$n. $($a['Name']) -> HTTP $($resp.StatusCode) $($resp.ContentType): $($read.Reason)")
        }
        $out.Raw = ($log -join "`n")

        if (-not $payload) {
            # The first attempt is the site's own request shape, so its answer
            # is the one worth repeating; the rest are in Raw.
            $first = [string]$log[0]
            if ($first -match 'not a parts list: (.+)$')     { $out.Error = "Lenovo refused the parts list: $($Matches[1])" }
            elseif ($first -match 'a web page')              { $out.Error = 'Lenovo returned a web page instead of a parts list' }
            elseif ($first -match '-> (Request failed.*)$')  { $out.Error = $Matches[1] }
            elseif ($first -match '-> HTTP (\d+)')           { $out.Error = "Lenovo did not return a parts list (HTTP $($Matches[1]))" }
            else                                             { $out.Error = 'Lenovo did not return a parts list' }
            return [pscustomobject]$out
        }

        $out.Source = $payload.Source
        $out.Parts  = @(ConvertTo-LenovoPartRows -Record $payload.Records)
        if ($out.Parts.Count -eq 0) { $out.Error = 'No parts listed for this serial' }
        else { $out.Raw = '' }
    } catch {
        $out.Error = "Could not read the parts list: $($_.Exception.Message)"
    } finally {
        $client.Dispose()
    }

    return [pscustomobject]$out
}


function Find-LenovoPartsEndpoint {
    <#
    .SYNOPSIS
        Diagnostic: scans a product's parts page and its scripts for the parts
        API the site itself calls.

    .DESCRIPTION
        Resolves the serial, fetches the product's parts page, then every
        script it loads from lenovo.com, and lists each api/v4 path mentioning
        parts, every mention of upsellAggregation with its surroundings, and
        whether the page HTML itself carries part numbers. Returns lines of
        text meant to be pasted into a bug report.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)] [string] $SerialNumber,
        [int] $TimeoutSec = 90,
        [int] $MaxScripts = 30
    )

    $lines  = New-Object System.Collections.Generic.List[string]
    $client = New-LenovoHttpClient -TimeoutSec $TimeoutSec

    $addHits = {
        param([string] $Text, [string] $Label)
        if (-not $Text) { return }
        # The query string is kept when the code spells one out, because the
        # parameter names are half of what needs knowing.
        $paths = @([regex]::Matches($Text, '(?i)/?api/v\d+/[A-Za-z0-9_./${}-]*part[A-Za-z0-9_./${}?&=-]*') |
                   ForEach-Object { $_.Value } | Sort-Object -Unique)
        foreach ($hit in $paths) { [void]$lines.Add("  [$Label] api path: $hit") }
        $ctx = @([regex]::Matches($Text, '.{0,120}upsellAggregation.{0,200}') | ForEach-Object { $_.Value } | Select-Object -First 25)
        foreach ($c in $ctx) { [void]$lines.Add("  [$Label] upsellAggregation context: " + (($c -replace '\s+', ' ').Trim())) }
    }

    try {
        $product = Get-LenovoProduct -SerialNumber $SerialNumber -Client $client
        if ($product.Error) {
            [void]$lines.Add("Product: $($product.Error) | $($product.Raw)")
            return $lines.ToArray()
        }
        [void]$lines.Add("Product: $($product.Product) | type $($product.MachineType) | model $($product.Model) | id $($product.ProductId)")

        $pageUrl = "$($script:SiteBase)/us/en/products/$($product.ProductId)/parts"
        $page    = Invoke-LenovoRequest -Client $client -Url $pageUrl -Accept 'text/html, */*'
        [void]$lines.Add("Parts page: $pageUrl -> HTTP $($page.StatusCode) $($page.ContentType), $($page.Bytes.Length) bytes $($page.Error)")
        if (-not $page.Text) { return $lines.ToArray() }

        # FRU numbers are 7 or 10 characters, start with a digit, mix in letters:
        # 01AV430, 45N0346, 5CB1H89763, 5M10Z39781.
        $frus = @([regex]::Matches($page.Text, '\b\d(?:[A-Z0-9]{6}|[A-Z0-9]{9})\b') |
                  ForEach-Object { $_.Value } | Where-Object { $_ -match '[A-Z]' } | Sort-Object -Unique)
        $eg   = if ($frus.Count -gt 0) { ' e.g. ' + (($frus | Select-Object -First 6) -join ', ') } else { '' }
        [void]$lines.Add("Part-number-looking tokens in the page HTML: $($frus.Count)$eg")
        & $addHits $page.Text 'page'

        $srcs = @([regex]::Matches($page.Text, '(?i)<script[^>]+src\s*=\s*["'']([^"'']+)["'']') |
                  ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        $base = [Uri]$pageUrl
        $urls = @()
        foreach ($src in $srcs) {
            $u = $null
            # Only the site's own scripts and Lenovo's CDNs; third party tags
            # (analytics, chat widgets) cannot know the parts API.
            if ([Uri]::TryCreate($base, $src, [ref]$u) -and
                ($u.Host -eq $base.Host -or $u.Host -match 'lenovo\.com$')) { $urls += $u.AbsoluteUri }
        }
        $urls = @($urls | Sort-Object -Unique | Select-Object -First $MaxScripts)
        [void]$lines.Add("Scripts on the page: $($srcs.Count), scanning $($urls.Count) from the site or lenovo.com")

        foreach ($u in $urls) {
            $js = Invoke-LenovoRequest -Client $client -Url $u -Accept '*/*'
            $name = ($u -split '/')[-1]
            if ($name.Length -gt 60) { $name = $name.Substring(0, 60) }
            $before = $lines.Count
            if ($js.Text) { & $addHits $js.Text $name }
            [void]$lines.Add("  script $name -> HTTP $($js.StatusCode), $($js.Bytes.Length) bytes, $($lines.Count - $before) hits")
        }
    } catch {
        [void]$lines.Add("Scan failed: $($_.Exception.Message)")
    } finally {
        $client.Dispose()
    }

    return $lines.ToArray()
}


function Find-LenovoPart {
    <#
    .SYNOPSIS
        Finds the part number(s) for one preset part on one Lenovo serial.

    .DESCRIPTION
        Looks up the serial's parts list and keeps the rows matching the
        chosen preset (see Get-LenovoPartCategory for the list). Returns the
        parts list object with Category and Matches added.

    .EXAMPLE
        (Find-LenovoPart PF0ABCDE 'System board').Matches
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)] [string] $SerialNumber,
        [Parameter(Position = 1)]            [string] $Part = 'All parts',
        [int] $TimeoutSec = 60
    )

    # Validate the preset before spending a round trip on it.
    if ($Part -and $Part -notmatch '^\s*all(\s*parts)?\s*$') { [void](Get-LenovoPartCategory -Name $Part) }

    $list = Get-LenovoPartsList -SerialNumber $SerialNumber -TimeoutSec $TimeoutSec
    $hits = @()
    if (-not $list.Error) { $hits = @(Select-LenovoPart -Part $list.Parts -Category $Part) }

    $list | Add-Member -NotePropertyName Category -NotePropertyValue $Part
    $list | Add-Member -NotePropertyName Matches  -NotePropertyValue $hits
    return $list
}


Export-ModuleMember -Function Get-LenovoWarranty, ConvertTo-SerialList,
                              Format-WarrantyDate, ConvertFrom-LenovoIbaseInfo,
                              ConvertTo-DateOrNull,
                              Get-LenovoProduct, Get-LenovoPartsList, Find-LenovoPart,
                              Get-LenovoPartCategory, Select-LenovoPart,
                              ConvertTo-LenovoPartRows, ConvertFrom-LenovoXlsx,
                              ConvertFrom-LenovoDelimited, ConvertFrom-LenovoPartsJson,
                              ConvertFrom-LenovoJsonList, Read-LenovoPartsPayload,
                              Find-LenovoPartsEndpoint
