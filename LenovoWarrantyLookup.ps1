#Requires -Version 5.1
<#
    LenovoWarrantyLookup.ps1

    Paste a batch of Lenovo serial numbers, get back the main device warranty
    end date for each one, in the same order, ready to paste into Excel.

    Launch it with "Lenovo Warranty Lookup.cmd", or run this file directly with
    powershell -ExecutionPolicy Bypass -File .\LenovoWarrantyLookup.ps1

    Visual treatment is deliberately brutalist: hard edges, heavy rules, solid
    offset shadows, monospace data, no gradients or rounded corners. The palette
    is taken from the project logo - crimson, magenta and near black navy.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $here 'LenovoWarranty.psm1') -Force


# --------------------------------------------------------------------------
#  Palette - straight off the logo
# --------------------------------------------------------------------------
$crimson = [System.Drawing.Color]::FromArgb(225, 11, 46)    # primary
$magenta = [System.Drawing.Color]::FromArgb(255, 15, 107)   # accent
$navy    = [System.Drawing.Color]::FromArgb(20, 20, 45)     # ink / rules
$bone    = [System.Drawing.Color]::FromArgb(245, 244, 240)  # page
$paper   = [System.Drawing.Color]::White
$ghost   = [System.Drawing.Color]::FromArgb(168, 168, 178)  # disabled

$BORDER  = 3    # rule weight
$DROP    = 5    # hard shadow offset

# Segoe UI Black is the heaviest face shipped with Windows; Consolas carries
# all the data so serials line up in a column the way they do in Excel.
$fontTitle = New-Object System.Drawing.Font('Segoe UI Black', 17)
$fontLabel = New-Object System.Drawing.Font('Consolas', 10, [System.Drawing.FontStyle]::Bold)
$fontBtn   = New-Object System.Drawing.Font('Consolas', 9.5, [System.Drawing.FontStyle]::Bold)
$fontMono  = New-Object System.Drawing.Font('Consolas', 10)
$fontHead  = New-Object System.Drawing.Font('Consolas', 9.5, [System.Drawing.FontStyle]::Bold)
$fontSub   = New-Object System.Drawing.Font('Consolas', 9)


# --------------------------------------------------------------------------
#  Helpers
# --------------------------------------------------------------------------
function New-BorderBox {
    <#  A solid navy panel used as a thick rule around whatever is docked in it. #>
    param($X, $Y, $W, $H, [string]$Anchor)

    $p           = New-Object System.Windows.Forms.Panel
    $p.Location  = New-Object System.Drawing.Point($X, $Y)
    $p.Size      = New-Object System.Drawing.Size($W, $H)
    $p.BackColor = $navy
    $p.Padding   = New-Object System.Windows.Forms.Padding($BORDER)
    if ($Anchor) { $p.Anchor = $Anchor }
    return $p
}

function New-BrutalButton {
    <#
        A flat block with a heavy rule and a solid offset shadow behind it.
        Returns the button; the shadow panel rides along in .Tag so the caller
        can add both and keep them in step.
    #>
    param(
        [string] $Text,
        [int]    $X,
        [int]    $Width,
        [System.Drawing.Color] $Fill,
        [System.Drawing.Color] $Ink
    )

    $shadow           = New-Object System.Windows.Forms.Panel
    $shadow.Location  = New-Object System.Drawing.Point(($X + $DROP), (10 + $DROP))
    $shadow.Size      = New-Object System.Drawing.Size($Width, 36)
    $shadow.BackColor = $navy

    $b           = New-Object System.Windows.Forms.Button
    $b.Text      = $Text
    $b.Location  = New-Object System.Drawing.Point($X, 10)
    $b.Size      = New-Object System.Drawing.Size($Width, 36)
    $b.FlatStyle = 'Flat'
    $b.Font      = $fontBtn
    $b.BackColor = $Fill
    $b.ForeColor = $Ink
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $b.FlatAppearance.BorderSize  = $BORDER
    $b.FlatAppearance.BorderColor = $navy
    $b.FlatAppearance.MouseOverBackColor = $magenta
    $b.FlatAppearance.MouseDownBackColor = $navy
    $b.TabStop   = $true

    # Remember the resting colours so enable/disable can restore them.
    $b.Tag = [pscustomobject]@{ Shadow = $shadow; Fill = $Fill; Ink = $Ink }

    # Pressing the block drives it into its own shadow.
    $b.Add_MouseDown({
        $this.Left += $DROP
        $this.Top  += $DROP
    })
    $b.Add_MouseUp({
        $this.Left -= $DROP
        $this.Top  -= $DROP
    })
    $b.Add_MouseEnter({ if ($this.Enabled) { $this.ForeColor = [System.Drawing.Color]::White } })
    $b.Add_MouseLeave({ if ($this.Enabled) { $this.ForeColor = $this.Tag.Ink } })

    return $b
}

function Set-BlockEnabled {
    <#  Flat buttons ignore their own disabled styling, so drive it by hand. #>
    param($Button, [bool] $On)

    $Button.Enabled = $On
    if ($On) {
        $Button.BackColor = $Button.Tag.Fill
        $Button.ForeColor = $Button.Tag.Ink
        $Button.FlatAppearance.BorderColor = $navy
        $Button.Tag.Shadow.BackColor = $navy
    } else {
        $Button.BackColor = $bone
        $Button.ForeColor = $ghost
        $Button.FlatAppearance.BorderColor = $ghost
        $Button.Tag.Shadow.BackColor = $bone
    }
}


# --------------------------------------------------------------------------
#  Window
# --------------------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'LENOVO BATCH WARRANTY LOOKUP'
$form.Size          = New-Object System.Drawing.Size(1120, 760)
$form.MinimumSize   = New-Object System.Drawing.Size(980, 640)
$form.StartPosition = 'CenterScreen'
$form.BackColor     = $bone
$form.Font          = $fontMono


# ---- masthead ----
$mast           = New-Object System.Windows.Forms.Panel
$mast.Location  = New-Object System.Drawing.Point(0, 0)
$mast.Size      = New-Object System.Drawing.Size(1120, 92)
$mast.Anchor    = 'Top,Left,Right'
$mast.BackColor = $navy
$form.Controls.Add($mast)

$mastTitle           = New-Object System.Windows.Forms.Label
$mastTitle.Text      = 'WARRANTY LOOKUP'
$mastTitle.Font      = $fontTitle
$mastTitle.ForeColor = $paper
$mastTitle.AutoSize  = $true
$mastTitle.BackColor = [System.Drawing.Color]::Transparent
$mastTitle.Location  = New-Object System.Drawing.Point(26, 18)
$mast.Controls.Add($mastTitle)

$mastSub           = New-Object System.Windows.Forms.Label
$mastSub.Text      = 'LENOVO // BATCH // MAIN DEVICE WARRANTY ONLY'
$mastSub.Font      = $fontSub
$mastSub.ForeColor = $magenta
$mastSub.AutoSize  = $true
$mastSub.BackColor = [System.Drawing.Color]::Transparent
$mastSub.Location  = New-Object System.Drawing.Point(29, 56)
$mast.Controls.Add($mastSub)

# Logo motif: offset bars in the two logo colours, echoing the mark.
$mast.Add_Paint({
    $g = $_.Graphics
    $g.SmoothingMode = 'None'
    $w = $mast.ClientSize.Width

    $bC = New-Object System.Drawing.SolidBrush($crimson)
    $bM = New-Object System.Drawing.SolidBrush($magenta)
    $bP = New-Object System.Drawing.SolidBrush($paper)

    # A loose scatter of thick strokes on an implied grid, the way the mark is
    # built: verticals and horizontals crossing, one small solid square adrift.
    $g.FillRectangle($bM, ($w - 156), 16, 13, 34)   # vertical, magenta
    $g.FillRectangle($bC, ($w - 138), 44, 34, 13)   # horizontal, crimson
    $g.FillRectangle($bC, ($w - 100), 14, 13, 26)   # vertical, crimson
    $g.FillRectangle($bM, ($w -  96), 50, 26, 13)   # horizontal, magenta
    $g.FillRectangle($bM, ($w -  62), 22, 13, 22)   # vertical, magenta
    $g.FillRectangle($bP, ($w -  40), 52, 11, 11)   # square, bone

    $bC.Dispose(); $bM.Dispose(); $bP.Dispose()
})

# Crimson rule under the masthead.
$mastRule           = New-Object System.Windows.Forms.Panel
$mastRule.Location  = New-Object System.Drawing.Point(0, 92)
$mastRule.Size      = New-Object System.Drawing.Size(1120, 6)
$mastRule.Anchor    = 'Top,Left,Right'
$mastRule.BackColor = $crimson
$form.Controls.Add($mastRule)


# --------------------------------------------------------------------------
#  Split: serials on the left, results on the right
# --------------------------------------------------------------------------
$split                  = New-Object System.Windows.Forms.SplitContainer
$split.Location         = New-Object System.Drawing.Point(22, 120)
$split.Size             = New-Object System.Drawing.Size(1064, 520)
$split.Anchor           = 'Top,Left,Right,Bottom'
$split.SplitterWidth    = 14
$split.BackColor        = $bone
$split.Panel1.BackColor = $bone
$split.Panel2.BackColor = $bone
$form.Controls.Add($split)
$split.SplitterDistance = 250

# ---- left ----
$lblIn           = New-Object System.Windows.Forms.Label
$lblIn.Text      = 'SERIAL NUMBERS'
$lblIn.Font      = $fontLabel
$lblIn.ForeColor = $navy
$lblIn.AutoSize  = $true
$lblIn.Location  = New-Object System.Drawing.Point(0, 0)
$split.Panel1.Controls.Add($lblIn)

$inBox = New-BorderBox -X 0 -Y 24 -W 244 -H 442 -Anchor 'Top,Left,Right,Bottom'
$split.Panel1.Controls.Add($inBox)

$txtIn               = New-Object System.Windows.Forms.TextBox
$txtIn.Multiline     = $true
$txtIn.ScrollBars    = 'Vertical'
$txtIn.AcceptsReturn = $true
$txtIn.WordWrap      = $false
$txtIn.BorderStyle   = 'None'
$txtIn.Font          = $fontMono
$txtIn.BackColor     = $paper
$txtIn.ForeColor     = $navy
$txtIn.Dock          = 'Fill'
$inBox.Controls.Add($txtIn)

$lblCount           = New-Object System.Windows.Forms.Label
$lblCount.Text      = '0 SERIALS'
$lblCount.Font      = $fontSub
$lblCount.ForeColor = $navy
$lblCount.AutoSize  = $true
$lblCount.Location  = New-Object System.Drawing.Point(0, 474)
$lblCount.Anchor    = 'Left,Bottom'
$split.Panel1.Controls.Add($lblCount)

# ---- right ----
$lblOut           = New-Object System.Windows.Forms.Label
$lblOut.Text      = 'RESULTS'
$lblOut.Font      = $fontLabel
$lblOut.ForeColor = $navy
$lblOut.AutoSize  = $true
$lblOut.Location  = New-Object System.Drawing.Point(0, 0)
$split.Panel2.Controls.Add($lblOut)

$gridBox = New-BorderBox -X 0 -Y 24 -W 796 -H 468 -Anchor 'Top,Left,Right,Bottom'
$split.Panel2.Controls.Add($gridBox)

$grid                           = New-Object System.Windows.Forms.DataGridView
$grid.Dock                      = 'Fill'
$grid.ReadOnly                  = $true
$grid.AllowUserToAddRows        = $false
$grid.AllowUserToDeleteRows     = $false
$grid.AllowUserToResizeRows     = $false
$grid.RowHeadersVisible         = $false
$grid.SelectionMode             = 'FullRowSelect'
$grid.MultiSelect               = $true
$grid.BorderStyle               = 'None'
$grid.BackgroundColor           = $paper
$grid.GridColor                 = $navy
$grid.CellBorderStyle           = 'Single'
$grid.ColumnHeadersBorderStyle  = 'Single'
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersHeightSizeMode = 'DisableResizing'
$grid.ColumnHeadersHeight       = 34
$grid.RowTemplate.Height        = 27

$grid.ColumnHeadersDefaultCellStyle.BackColor = $navy
$grid.ColumnHeadersDefaultCellStyle.ForeColor = $paper
$grid.ColumnHeadersDefaultCellStyle.Font      = $fontHead
$grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $navy
$grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = $paper
$grid.ColumnHeadersDefaultCellStyle.Padding   = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)

$grid.DefaultCellStyle.Font               = $fontMono
$grid.DefaultCellStyle.ForeColor          = $navy
$grid.DefaultCellStyle.BackColor          = $paper
$grid.DefaultCellStyle.SelectionBackColor = $magenta
$grid.DefaultCellStyle.SelectionForeColor = $paper
$grid.DefaultCellStyle.Padding            = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
$grid.AlternatingRowsDefaultCellStyle.BackColor = $bone
$grid.AlternatingRowsDefaultCellStyle.SelectionBackColor = $magenta
$gridBox.Controls.Add($grid)

foreach ($spec in @(
    @{ N = 'Serial';      H = 'SERIAL';   W = 112 },
    @{ N = 'WarrantyEnd'; H = 'ENDS';     W = 114 },
    @{ N = 'Status';      H = 'STATUS';   W = 92  },
    @{ N = 'Product';     H = 'PRODUCT';  W = 240 },
    @{ N = 'Note';        H = 'COVERAGE'; W = 190 }
)) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name       = $spec.N
    $col.HeaderText = $spec.H
    $col.Width      = $spec.W
    [void]$grid.Columns.Add($col)
}
$grid.Columns['Product'].AutoSizeMode = 'Fill'


# --------------------------------------------------------------------------
#  Control bar
# --------------------------------------------------------------------------
$barRule           = New-Object System.Windows.Forms.Panel
$barRule.Location  = New-Object System.Drawing.Point(22, 652)
$barRule.Size      = New-Object System.Drawing.Size(1064, $BORDER)
$barRule.Anchor    = 'Left,Right,Bottom'
$barRule.BackColor = $navy
$form.Controls.Add($barRule)

$bar           = New-Object System.Windows.Forms.Panel
$bar.Location  = New-Object System.Drawing.Point(22, 662)
$bar.Size      = New-Object System.Drawing.Size(1064, 58)
$bar.Anchor    = 'Left,Right,Bottom'
$bar.BackColor = $bone
$form.Controls.Add($bar)

$btnLookup    = New-BrutalButton -Text 'LOOK UP'    -X 0   -Width 132 -Fill $crimson -Ink $paper
$btnCopyDates = New-BrutalButton -Text 'COPY DATES' -X 146 -Width 132 -Fill $magenta -Ink $paper
$btnCopyTable = New-BrutalButton -Text 'COPY TABLE' -X 292 -Width 132 -Fill $paper   -Ink $navy
$btnSaveCsv   = New-BrutalButton -Text 'SAVE CSV'   -X 438 -Width 118 -Fill $paper   -Ink $navy
$btnClear     = New-BrutalButton -Text 'CLEAR'      -X 570 -Width  94 -Fill $paper   -Ink $navy

foreach ($b in @($btnLookup, $btnCopyDates, $btnCopyTable, $btnSaveCsv, $btnClear)) {
    $bar.Controls.Add($b.Tag.Shadow)
    $bar.Controls.Add($b)
    $b.BringToFront()
}

$lblFmt           = New-Object System.Windows.Forms.Label
$lblFmt.Text      = 'FORMAT'
$lblFmt.Font      = $fontSub
$lblFmt.ForeColor = $navy
$lblFmt.AutoSize  = $true
$lblFmt.Location  = New-Object System.Drawing.Point(686, 21)
$bar.Controls.Add($lblFmt)

$fmtBox = New-BorderBox -X 744 -Y 10 -W 132 -H 30
$bar.Controls.Add($fmtBox)

$cboFmt                = New-Object System.Windows.Forms.ComboBox
$cboFmt.DropDownStyle  = 'DropDownList'
$cboFmt.FlatStyle      = 'Flat'
$cboFmt.Font           = $fontBtn
$cboFmt.BackColor      = $paper
$cboFmt.ForeColor      = $navy
$cboFmt.Dock           = 'Fill'
[void]$cboFmt.Items.AddRange(@('MM/dd/yyyy', 'yyyy-MM-dd', 'dd/MM/yyyy', 'M/d/yyyy'))
$cboFmt.SelectedIndex  = 0
$fmtBox.Controls.Add($cboFmt)

# Blocky progress meter, drawn rather than themed.
$progShell           = New-Object System.Windows.Forms.Panel
$progShell.Location  = New-Object System.Drawing.Point(896, 10)
$progShell.Size      = New-Object System.Drawing.Size(168, 30)
$progShell.Anchor    = 'Right,Bottom'
$progShell.BackColor = $navy
$progShell.Padding   = New-Object System.Windows.Forms.Padding($BORDER)
$progShell.Visible   = $false
$bar.Controls.Add($progShell)

$progTrack           = New-Object System.Windows.Forms.Panel
$progTrack.Dock      = 'Fill'
$progTrack.BackColor = $paper
$progShell.Controls.Add($progTrack)

$progFill            = New-Object System.Windows.Forms.Panel
$progFill.Location   = New-Object System.Drawing.Point(0, 0)
$progFill.Size       = New-Object System.Drawing.Size(0, 24)
$progFill.BackColor  = $crimson
$progTrack.Controls.Add($progFill)

$status           = New-Object System.Windows.Forms.Label
$status.Text      = 'READY'
$status.Font      = $fontLabel
$status.ForeColor = $navy
$status.AutoSize  = $true
$status.Location  = New-Object System.Drawing.Point(896, 20)
$status.Anchor    = 'Right,Bottom'
$bar.Controls.Add($status)


# --------------------------------------------------------------------------
#  Behaviour
# --------------------------------------------------------------------------
$script:Results = @()

function Get-CurrentFormat { return [string]$cboFmt.SelectedItem }

function Update-InputCount {
    $n = @(ConvertTo-SerialList -Text $txtIn.Text).Count
    $lblCount.Text = if ($n -eq 1) { '1 SERIAL' } else { "$n SERIALS" }
}

function Set-Status {
    param([string]$Text, [System.Drawing.Color]$Color = $navy)
    $status.Text      = $Text.ToUpperInvariant()
    $status.ForeColor = $Color
    $status.Visible   = $true
}

function Set-Progress {
    param([int]$Done, [int]$Total)
    if ($Total -le 0) { return }
    $w = [int]($progTrack.ClientSize.Width * $Done / $Total)
    $progFill.Size = New-Object System.Drawing.Size($w, $progTrack.ClientSize.Height)
}

function Update-Grid {
    <#  Repaints the grid from $script:Results using the chosen date format. #>
    $grid.SuspendLayout()
    $grid.Rows.Clear()

    foreach ($r in $script:Results) {
        $dateText = Format-WarrantyDate -Date $r.WarrantyEnd -Format (Get-CurrentFormat)
        $note     = if ($r.Error) { $r.Error.ToUpperInvariant() } else { $r.WarrantyName }

        # Terse wording keeps the column narrow and suits the rest of the type.
        $statusText = switch ($r.Status) {
            'In warranty'     { 'ACTIVE' }
            'Out of warranty' { 'EXPIRED' }
            default           { $r.Status.ToUpperInvariant() }
        }

        $idx = $grid.Rows.Add($r.Serial, $dateText, $statusText, $r.Product, $note)
        $row = $grid.Rows[$idx]

        # The date is the whole point of the tool, so it gets the loudest colour.
        $row.Cells['WarrantyEnd'].Style.Font      = $fontLabel
        $row.Cells['WarrantyEnd'].Style.ForeColor = $crimson
        $row.Cells['Serial'].Style.Font           = $fontLabel

        # Live coverage reads as a solid block rather than just coloured text.
        if ($statusText -eq 'ACTIVE') {
            $row.Cells['Status'].Style.BackColor = $crimson
            $row.Cells['Status'].Style.ForeColor = $paper
            $row.Cells['Status'].Style.Font      = $fontLabel
        }

        if ($r.Error) {
            $row.Cells['Serial'].Style.ForeColor = $magenta
            $row.Cells['Note'].Style.ForeColor   = $magenta
            $row.Cells['Note'].Style.Font        = $fontLabel
        }
    }

    $grid.ResumeLayout()

    # Nothing is "chosen" yet, and a full width magenta bar over row one just
    # hides the first result.
    $grid.ClearSelection()
}

function Get-DatesText {
    <#  One date per line, matching the input order - a single Excel column. #>
    $fmt = Get-CurrentFormat
    return (($script:Results | ForEach-Object {
        Format-WarrantyDate -Date $_.WarrantyEnd -Format $fmt
    }) -join "`r`n")
}

function Get-TableText {
    <#  Serial + date + status, tab separated, with a header row for Excel. #>
    $fmt   = Get-CurrentFormat
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("Serial`tWarranty End`tStatus`tProduct`tCoverage")

    foreach ($r in $script:Results) {
        $dateText = Format-WarrantyDate -Date $r.WarrantyEnd -Format $fmt
        $note     = if ($r.Error) { $r.Error } else { $r.WarrantyName }
        [void]$lines.Add(("{0}`t{1}`t{2}`t{3}`t{4}" -f $r.Serial, $dateText, $r.Status, $r.Product, $note))
    }

    return ($lines -join "`r`n")
}

function Copy-ToClipboardSafe {
    <#
        The clipboard is shared with every other app, so a lock by Excel or a
        remote session can make the first attempt fail. Retry briefly instead
        of throwing a wall of red at the user.
    #>
    param([string]$Text, [string]$Label)

    if ([string]::IsNullOrEmpty($Text)) {
        Set-Status 'NOTHING TO COPY' $magenta
        return
    }

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            [System.Windows.Forms.Clipboard]::SetText($Text)
            Set-Status "$Label COPIED" $crimson
            return
        } catch {
            Start-Sleep -Milliseconds 120
        }
    }
    Set-Status 'CLIPBOARD BLOCKED' $magenta
}

function Invoke-Lookup {
    $serials = @(ConvertTo-SerialList -Text $txtIn.Text)

    if ($serials.Count -eq 0) {
        Set-Status 'PASTE SERIALS FIRST' $magenta
        return
    }

    foreach ($b in @($btnLookup, $btnClear, $btnCopyDates, $btnCopyTable, $btnSaveCsv)) {
        Set-BlockEnabled $b $false
    }

    $status.Visible   = $false
    $progShell.Visible = $true
    $progFill.Size    = New-Object System.Drawing.Size(0, $progTrack.ClientSize.Height)
    $form.Cursor      = [System.Windows.Forms.Cursors]::WaitCursor

    # Each completed request advances the meter and lets the window repaint, so
    # the UI stays alive through a long batch.
    $tick = {
        param($doneCount, $totalCount)
        Set-Progress $doneCount $totalCount
        [System.Windows.Forms.Application]::DoEvents()
    }

    try {
        $script:Results = @(Get-LenovoWarranty -SerialNumber $serials -OnProgress $tick)
        Update-Grid

        $found  = @($script:Results | Where-Object { $_.WarrantyEnd }).Count
        $failed = $script:Results.Count - $found

        foreach ($b in @($btnCopyDates, $btnCopyTable, $btnSaveCsv)) {
            Set-BlockEnabled $b $true
        }

        if ($failed -gt 0) {
            Set-Status "$found/$($script:Results.Count) FOUND - $failed MISSED" $magenta
        } else {
            Set-Status "$found/$found FOUND" $crimson
        }
    } catch {
        Set-Status "FAILED: $($_.Exception.Message)" $magenta
    } finally {
        $progShell.Visible = $false
        $form.Cursor       = [System.Windows.Forms.Cursors]::Default
        Set-BlockEnabled $btnLookup $true
        Set-BlockEnabled $btnClear  $true
    }
}


# --------------------------------------------------------------------------
#  Wiring
# --------------------------------------------------------------------------
$txtIn.Add_TextChanged({ Update-InputCount })

$btnLookup.Add_Click({ Invoke-Lookup })

$btnCopyDates.Add_Click({ Copy-ToClipboardSafe -Text (Get-DatesText) -Label 'DATES' })

$btnCopyTable.Add_Click({ Copy-ToClipboardSafe -Text (Get-TableText) -Label 'TABLE' })

$btnClear.Add_Click({
    $txtIn.Clear()
    $grid.Rows.Clear()
    $script:Results = @()
    foreach ($b in @($btnCopyDates, $btnCopyTable, $btnSaveCsv)) { Set-BlockEnabled $b $false }
    Set-Status 'READY'
    $txtIn.Focus()
})

$cboFmt.Add_SelectedIndexChanged({
    if ($script:Results.Count -gt 0) { Update-Grid }
})

$btnSaveCsv.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'CSV file (*.csv)|*.csv'
    $dlg.FileName = "Lenovo warranty $(Get-Date -Format 'yyyy-MM-dd').csv"

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $fmt = Get-CurrentFormat
            $script:Results | ForEach-Object {
                [pscustomobject]@{
                    Serial       = $_.Serial
                    WarrantyEnd  = Format-WarrantyDate -Date $_.WarrantyEnd -Format $fmt
                    Status       = $_.Status
                    Product      = $_.Product
                    MachineType  = $_.MachineType
                    Coverage     = $_.WarrantyName
                    Note         = $_.Error
                }
            } | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8

            Set-Status 'CSV SAVED' $crimson
        } catch {
            Set-Status "SAVE FAILED: $($_.Exception.Message)" $magenta
        }
    }
})

# Ctrl+Enter runs the lookup from the serial box.
$txtIn.Add_KeyDown({
    if ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $_.SuppressKeyPress = $true
        Invoke-Lookup
    }
})

$form.Add_Shown({
    foreach ($b in @($btnCopyDates, $btnCopyTable, $btnSaveCsv)) { Set-BlockEnabled $b $false }
    $txtIn.Focus()
})

[void]$form.ShowDialog()
$form.Dispose()
