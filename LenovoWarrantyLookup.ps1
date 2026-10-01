#Requires -Version 5.1
<#
    LenovoWarrantyLookup.ps1

    Paste a batch of Lenovo serial numbers, get back the main device warranty
    end date for each one, in the same order, ready to paste into Excel.

    The PARTS tab in the masthead switches to a second section: type one
    serial, pick a part from the preset list (LCD back cover, SSD, system
    board, ...) and it returns the Lenovo part number(s) for that machine.

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

function New-BrutalGrid {
    <#
        A results grid in the house style: navy header block, hard single
        rules, bone zebra rows, magenta selection. Columns come in as
        @{ N = name; H = header; W = width } and one of them can stretch.
    #>
    param([object[]] $Columns, [string] $FillColumn)

    $g                           = New-Object System.Windows.Forms.DataGridView
    $g.Dock                      = 'Fill'
    $g.ReadOnly                  = $true
    $g.AllowUserToAddRows        = $false
    $g.AllowUserToDeleteRows     = $false
    $g.AllowUserToResizeRows     = $false
    $g.RowHeadersVisible         = $false
    $g.SelectionMode             = 'FullRowSelect'
    $g.MultiSelect               = $true
    $g.BorderStyle               = 'None'
    $g.BackgroundColor           = $paper
    $g.GridColor                 = $navy
    $g.CellBorderStyle           = 'Single'
    $g.ColumnHeadersBorderStyle  = 'Single'
    $g.EnableHeadersVisualStyles = $false
    $g.ColumnHeadersHeightSizeMode = 'DisableResizing'
    $g.ColumnHeadersHeight       = 34
    $g.RowTemplate.Height        = 27

    $g.ColumnHeadersDefaultCellStyle.BackColor = $navy
    $g.ColumnHeadersDefaultCellStyle.ForeColor = $paper
    $g.ColumnHeadersDefaultCellStyle.Font      = $fontHead
    $g.ColumnHeadersDefaultCellStyle.SelectionBackColor = $navy
    $g.ColumnHeadersDefaultCellStyle.SelectionForeColor = $paper
    $g.ColumnHeadersDefaultCellStyle.Padding   = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)

    $g.DefaultCellStyle.Font               = $fontMono
    $g.DefaultCellStyle.ForeColor          = $navy
    $g.DefaultCellStyle.BackColor          = $paper
    $g.DefaultCellStyle.SelectionBackColor = $magenta
    $g.DefaultCellStyle.SelectionForeColor = $paper
    $g.DefaultCellStyle.Padding            = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
    $g.AlternatingRowsDefaultCellStyle.BackColor = $bone
    $g.AlternatingRowsDefaultCellStyle.SelectionBackColor = $magenta

    foreach ($spec in $Columns) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name       = $spec.N
        $col.HeaderText = $spec.H
        $col.Width      = $spec.W
        [void]$g.Columns.Add($col)
    }
    if ($FillColumn) { $g.Columns[$FillColumn].AutoSizeMode = 'Fill' }

    return $g
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

# Author credit, hugging the right edge of the masthead where the placeholder
# mark used to sit. Repositioned on resize since the masthead stretches with
# the window (Anchor alone won't hug the right edge for an AutoSize label).
$mastCredit           = New-Object System.Windows.Forms.Label
$mastCredit.Text      = 'CREATED BY AIDEN ORTEGA'
$mastCredit.Font      = $fontSub
$mastCredit.ForeColor = $magenta
$mastCredit.AutoSize  = $true
$mastCredit.BackColor = [System.Drawing.Color]::Transparent
$mast.Controls.Add($mastCredit)

$positionCredit = {
    $mastCredit.Location = New-Object System.Drawing.Point(
        ($mast.ClientSize.Width - $mastCredit.PreferredSize.Width - 26), 40)
}
$mast.Add_Resize($positionCredit)
& $positionCredit

# Section switch, sitting between the title and the credit. Two blocks in
# the masthead's own ink: the live one is knocked out to paper.
function New-ModeTab {
    param([string] $Text, [int] $X, [int] $Width)

    $b           = New-Object System.Windows.Forms.Button
    $b.Text      = $Text
    $b.Location  = New-Object System.Drawing.Point($X, 28)
    $b.Size      = New-Object System.Drawing.Size($Width, 36)
    $b.FlatStyle = 'Flat'
    $b.Font      = $fontBtn
    $b.BackColor = $navy
    $b.ForeColor = $paper
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $b.TabStop   = $false
    $b.FlatAppearance.BorderSize  = $BORDER
    $b.FlatAppearance.BorderColor = $paper
    $b.FlatAppearance.MouseOverBackColor = $magenta
    $b.FlatAppearance.MouseDownBackColor = $crimson
    return $b
}

$tabWarranty = New-ModeTab -Text 'WARRANTY' -X 420 -Width 132
$tabParts    = New-ModeTab -Text 'PARTS'    -X 566 -Width 110
$mast.Controls.Add($tabWarranty)
$mast.Controls.Add($tabParts)

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

$grid = New-BrutalGrid -FillColumn 'Product' -Columns @(
    @{ N = 'Serial';      H = 'SERIAL';   W = 112 },
    @{ N = 'WarrantyEnd'; H = 'ENDS';     W = 114 },
    @{ N = 'Status';      H = 'STATUS';   W = 92  },
    @{ N = 'Product';     H = 'PRODUCT';  W = 240 },
    @{ N = 'Note';        H = 'COVERAGE'; W = 190 }
)
$gridBox.Controls.Add($grid)


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
#  Parts section - one serial, one part, its part number(s)
#
#  Sits exactly where the warranty split and bar sit and is swapped in by
#  the PARTS tab. Inputs down the left, part numbers on the right.
# --------------------------------------------------------------------------
$partsPanel           = New-Object System.Windows.Forms.Panel
$partsPanel.Location  = New-Object System.Drawing.Point(22, 120)
$partsPanel.Size      = New-Object System.Drawing.Size(1064, 520)
$partsPanel.Anchor    = 'Top,Left,Right,Bottom'
$partsPanel.BackColor = $bone
$partsPanel.Visible   = $false
$form.Controls.Add($partsPanel)

$PCOL = 300    # width of the input column

# ---- serial ----
$lblPSerial           = New-Object System.Windows.Forms.Label
$lblPSerial.Text      = 'SERIAL NUMBER'
$lblPSerial.Font      = $fontLabel
$lblPSerial.ForeColor = $navy
$lblPSerial.AutoSize  = $true
$lblPSerial.Location  = New-Object System.Drawing.Point(0, 0)
$partsPanel.Controls.Add($lblPSerial)

$pSerialBox = New-BorderBox -X 0 -Y 24 -W $PCOL -H 30
$partsPanel.Controls.Add($pSerialBox)

# Multiline so it fills its rule box; Enter is caught below and runs the
# lookup instead of adding a line.
$txtPSerial                = New-Object System.Windows.Forms.TextBox
$txtPSerial.Multiline      = $true
$txtPSerial.WordWrap       = $false
$txtPSerial.BorderStyle    = 'None'
$txtPSerial.Font           = $fontLabel
$txtPSerial.BackColor      = $paper
$txtPSerial.ForeColor      = $navy
$txtPSerial.CharacterCasing = 'Upper'
$txtPSerial.Dock           = 'Fill'
$pSerialBox.Controls.Add($txtPSerial)

# ---- part ----
$lblPPart           = New-Object System.Windows.Forms.Label
$lblPPart.Text      = 'PART'
$lblPPart.Font      = $fontLabel
$lblPPart.ForeColor = $navy
$lblPPart.AutoSize  = $true
$lblPPart.Location  = New-Object System.Drawing.Point(0, 70)
$partsPanel.Controls.Add($lblPPart)

$pPartBox = New-BorderBox -X 0 -Y 94 -W $PCOL -H 30
$partsPanel.Controls.Add($pPartBox)

# The preset list comes from the module so the GUI and the command line
# always agree on what can be asked for.
$script:PartCategoryNames = @(Get-LenovoPartCategory | ForEach-Object { $_.Name })

$cboPart                = New-Object System.Windows.Forms.ComboBox
$cboPart.DropDownStyle  = 'DropDownList'
$cboPart.FlatStyle      = 'Flat'
$cboPart.Font           = $fontBtn
$cboPart.BackColor      = $paper
$cboPart.ForeColor      = $navy
$cboPart.Dock           = 'Fill'
$cboPart.MaxDropDownItems = 16
[void]$cboPart.Items.Add('ALL PARTS')
foreach ($n in $script:PartCategoryNames) { [void]$cboPart.Items.Add($n.ToUpperInvariant()) }
$cboPart.SelectedIndex  = 0
$pPartBox.Controls.Add($cboPart)

# ---- product card ----
$lblPProduct           = New-Object System.Windows.Forms.Label
$lblPProduct.Text      = 'PRODUCT'
$lblPProduct.Font      = $fontLabel
$lblPProduct.ForeColor = $navy
$lblPProduct.AutoSize  = $true
$lblPProduct.Location  = New-Object System.Drawing.Point(0, 140)
$partsPanel.Controls.Add($lblPProduct)

$pProductBox = New-BorderBox -X 0 -Y 164 -W $PCOL -H 128
$partsPanel.Controls.Add($pProductBox)

$lblPInfo           = New-Object System.Windows.Forms.Label
$lblPInfo.Text      = ''
$lblPInfo.Font      = $fontSub
$lblPInfo.ForeColor = $navy
$lblPInfo.BackColor = $paper
$lblPInfo.AutoSize  = $false
$lblPInfo.Dock      = 'Fill'
$lblPInfo.Padding   = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
$lblPInfo.TextAlign = 'TopLeft'
$pProductBox.Controls.Add($lblPInfo)

$lblPHint           = New-Object System.Windows.Forms.Label
$lblPHint.Text      = "TYPE A SERIAL, PICK A PART, PRESS ENTER.`r`n`r`n" +
                      "MATCHES GO BY LENOVO'S OWN WORDING -`r`n" +
                      "READ THE DESCRIPTION BEFORE ORDERING.`r`n`r`n" +
                      "ALL PARTS SHOWS THE WHOLE LIST."
$lblPHint.Font      = $fontSub
$lblPHint.ForeColor = $navy
$lblPHint.AutoSize  = $true
$lblPHint.Location  = New-Object System.Drawing.Point(0, 306)
$partsPanel.Controls.Add($lblPHint)

# ---- part numbers ----
$lblPOut           = New-Object System.Windows.Forms.Label
$lblPOut.Text      = 'PART NUMBERS'
$lblPOut.Font      = $fontLabel
$lblPOut.ForeColor = $navy
$lblPOut.AutoSize  = $true
$lblPOut.Location  = New-Object System.Drawing.Point(($PCOL + 24), 0)
$partsPanel.Controls.Add($lblPOut)

$pGridBox = New-BorderBox -X ($PCOL + 24) -Y 24 -W (1064 - $PCOL - 24) -H 496 -Anchor 'Top,Left,Right,Bottom'
$partsPanel.Controls.Add($pGridBox)

$pgrid = New-BrutalGrid -FillColumn 'Description' -Columns @(
    @{ N = 'PartNumber';  H = 'PART NO.';    W = 132 },
    @{ N = 'Description'; H = 'DESCRIPTION'; W = 240 },
    @{ N = 'Commodity';   H = 'COMMODITY';   W = 150 },
    @{ N = 'Status';      H = 'STATUS';      W = 104 },
    @{ N = 'Substitutes'; H = 'SUBSTITUTES'; W = 120 }
)
$pGridBox.Controls.Add($pgrid)

# ---- parts control bar, over the top of the warranty one ----
$pbar           = New-Object System.Windows.Forms.Panel
$pbar.Location  = New-Object System.Drawing.Point(22, 662)
$pbar.Size      = New-Object System.Drawing.Size(1064, 58)
$pbar.Anchor    = 'Left,Right,Bottom'
$pbar.BackColor = $bone
$pbar.Visible   = $false
$form.Controls.Add($pbar)

$btnFind       = New-BrutalButton -Text 'FIND PART'     -X 0   -Width 132 -Fill $crimson -Ink $paper
$btnCopyPart   = New-BrutalButton -Text 'COPY PART NO.' -X 146 -Width 152 -Fill $magenta -Ink $paper
$btnCopyPTable = New-BrutalButton -Text 'COPY TABLE'    -X 312 -Width 132 -Fill $paper   -Ink $navy
$btnPClear     = New-BrutalButton -Text 'CLEAR'         -X 458 -Width  94 -Fill $paper   -Ink $navy

foreach ($b in @($btnFind, $btnCopyPart, $btnCopyPTable, $btnPClear)) {
    $pbar.Controls.Add($b.Tag.Shadow)
    $pbar.Controls.Add($b)
    $b.BringToFront()
}

$pstatus           = New-Object System.Windows.Forms.Label
$pstatus.Text      = 'READY'
$pstatus.Font      = $fontLabel
$pstatus.ForeColor = $navy
$pstatus.AutoSize  = $true
$pstatus.Location  = New-Object System.Drawing.Point(896, 20)
$pstatus.Anchor    = 'Right,Bottom'
$pbar.Controls.Add($pstatus)


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
    <#
        Writes to the warranty bar unless told which status label to use.
        Long messages are cut, and the label is kept flush with the right
        edge so a longer one grows leftwards rather than off the window.
    #>
    param([string]$Text, [System.Drawing.Color]$Color = $navy, $Target = $status)
    $t = $Text.ToUpperInvariant()
    if ($t.Length -gt 44) { $t = $t.Substring(0, 43) + '~' }
    $Target.Text      = $t
    $Target.ForeColor = $Color
    $Target.Visible   = $true
    $Target.Left      = $Target.Parent.ClientSize.Width - $Target.PreferredSize.Width
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
    param([string]$Text, [string]$Label, $Target = $status)

    if ([string]::IsNullOrEmpty($Text)) {
        Set-Status 'NOTHING TO COPY' $magenta -Target $Target
        return
    }

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            [System.Windows.Forms.Clipboard]::SetText($Text)
            Set-Status "$Label COPIED" $crimson -Target $Target
            return
        } catch {
            Start-Sleep -Milliseconds 120
        }
    }
    Set-Status 'CLIPBOARD BLOCKED' $magenta -Target $Target
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


# ---- parts section ----
$script:PartsList   = $null    # last Get-LenovoPartsList result, one serial
$script:PartMatches = @()      # rows currently in the parts grid

function Get-SelectedPartCategory {
    if ($cboPart.SelectedIndex -le 0) { return 'All parts' }
    return $script:PartCategoryNames[$cboPart.SelectedIndex - 1]
}

function Update-PartInfo {
    <#  The product card: what the serial resolved to, or why it did not. #>
    if (-not $script:PartsList) {
        $lblPInfo.Text = ''
        return
    }
    $r = $script:PartsList
    if ($r.Error) {
        $lblPInfo.ForeColor = $magenta
        $text = "$($r.Serial)`r`n`r`n$($r.Error.ToUpperInvariant())"
        if ($r.PSObject.Properties['Raw'] -and $r.Raw) {
            # What Lenovo actually sent, so a wrong guess about the endpoint
            # is visible in the window rather than silent. The command line's
            # -Diagnose switch has the full story.
            $raw = ([string]$r.Raw -split "`n")[0] -replace '^\d+\. ', ''
            if ($raw.Length -gt 150) { $raw = $raw.Substring(0, 150) + '...' }
            $text += "`r`n`r`nLENOVO SAID: $raw"
            $text += "`r`n`r`nRUN LOOKUP-PART.PS1 <SERIAL> -DIAGNOSE FOR DETAILS"
        }
        $lblPInfo.Text = $text
        return
    }
    $lblPInfo.ForeColor = $navy
    $lblPInfo.Text      = "$($r.Product)`r`n`r`nTYPE $($r.MachineType)   MODEL $($r.Model)`r`n$($r.Parts.Count) PARTS LISTED"
}

function Update-PartGrid {
    <#  Filters the cached parts list by the chosen preset and repaints. #>
    $cat = Get-SelectedPartCategory

    $script:PartMatches = @()
    if ($script:PartsList -and -not $script:PartsList.Error) {
        $script:PartMatches = @(Select-LenovoPart -Part $script:PartsList.Parts -Category $cat)
    }

    $pgrid.SuspendLayout()
    $pgrid.Rows.Clear()
    foreach ($p in $script:PartMatches) {
        $statusText = $p.Status.ToUpperInvariant()
        $idx = $pgrid.Rows.Add($p.PartNumber, $p.Description, $p.Commodity, $statusText, $p.Substitutes)
        $row = $pgrid.Rows[$idx]

        # The part number is what you came for, so it gets the loud colour.
        $row.Cells['PartNumber'].Style.Font      = $fontLabel
        $row.Cells['PartNumber'].Style.ForeColor = $crimson

        if ($statusText -match 'UNAVAIL|NOT AVAIL|DISCONTIN|NO LONGER|OBSOLETE|END OF LIFE|\bEOL\b') {
            $row.Cells['Status'].Style.ForeColor = $magenta
            $row.Cells['Status'].Style.Font      = $fontLabel
        }
    }
    $pgrid.ResumeLayout()
    $pgrid.ClearSelection()

    Update-PartInfo

    $have = $script:PartMatches.Count -gt 0
    foreach ($b in @($btnCopyPart, $btnCopyPTable)) { Set-BlockEnabled $b $have }

    if (-not $script:PartsList) { return }
    if ($script:PartsList.Error) {
        Set-Status $script:PartsList.Error $magenta -Target $pstatus
        return
    }

    $n = $script:PartMatches.Count
    if ($cat -eq 'All parts') {
        Set-Status "$n PARTS LISTED" $crimson -Target $pstatus
    } elseif ($have) {
        $word = if ($n -eq 1) { 'MATCH' } else { 'MATCHES' }
        Set-Status "$n $word FOR $cat" $crimson -Target $pstatus
    } else {
        Set-Status "NO $cat LISTED - TRY ALL PARTS" $magenta -Target $pstatus
    }
}

function Invoke-PartLookup {
    $serials = @(ConvertTo-SerialList -Text $txtPSerial.Text)
    if ($serials.Count -eq 0) {
        Set-Status 'ENTER A SERIAL FIRST' $magenta -Target $pstatus
        return
    }
    $serial = $serials[0]
    if ($serials.Count -gt 1) { $txtPSerial.Text = $serial }   # one machine at a time here

    # The list is per serial, so asking for another part on the same serial
    # is a re-filter of what was already downloaded, not another round trip.
    if ($script:PartsList -and $script:PartsList.Serial -eq $serial -and -not $script:PartsList.Error) {
        Update-PartGrid
        return
    }

    foreach ($b in @($btnFind, $btnPClear, $btnCopyPart, $btnCopyPTable)) { Set-BlockEnabled $b $false }
    Set-Status 'LOOKING UP' $navy -Target $pstatus
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $script:PartsList = Get-LenovoPartsList -SerialNumber $serial
        Update-PartGrid
    } catch {
        $script:PartsList = $null
        $pgrid.Rows.Clear()
        Set-Status "FAILED: $($_.Exception.Message)" $magenta -Target $pstatus
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        Set-BlockEnabled $btnFind   $true
        Set-BlockEnabled $btnPClear $true
    }
}

function Get-PartNumbersText {
    <#  Selected rows if any are picked, otherwise every match - one per line. #>
    $picked = @($pgrid.SelectedRows | Sort-Object Index)
    $nums = if ($picked.Count -gt 0) {
        $picked | ForEach-Object { [string]$_.Cells['PartNumber'].Value }
    } else {
        $script:PartMatches | ForEach-Object { $_.PartNumber }
    }
    return (@($nums) -join "`r`n")
}

function Get-PartTableText {
    <#  Every match with the serial and product it belongs to, for Excel. #>
    if (-not $script:PartsList) { return '' }
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("Serial`tProduct`tPart Number`tDescription`tCommodity`tStatus`tSubstitutes")
    foreach ($p in $script:PartMatches) {
        [void]$lines.Add(("{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}" -f
            $script:PartsList.Serial, $script:PartsList.Product,
            $p.PartNumber, $p.Description, $p.Commodity, $p.Status, $p.Substitutes))
    }
    return ($lines -join "`r`n")
}

function Set-Mode {
    <#  Swaps the warranty and parts sections and restyles the masthead tabs. #>
    param([ValidateSet('Warranty', 'Parts')] [string] $Mode)

    $parts = ($Mode -eq 'Parts')
    $split.Visible      = -not $parts
    $bar.Visible        = -not $parts
    $partsPanel.Visible = $parts
    $pbar.Visible       = $parts

    foreach ($pair in @(@($tabWarranty, (-not $parts)), @($tabParts, $parts))) {
        $tab = $pair[0]
        if ($pair[1]) { $tab.BackColor = $paper; $tab.ForeColor = $navy }
        else          { $tab.BackColor = $navy;  $tab.ForeColor = $paper }
    }

    if ($parts) {
        $mastTitle.Text = 'PART LOOKUP'
        $mastSub.Text   = 'LENOVO // ONE SERIAL // FRU PART NUMBERS'
        $txtPSerial.Focus()
    } else {
        $mastTitle.Text = 'WARRANTY LOOKUP'
        $mastSub.Text   = 'LENOVO // BATCH // MAIN DEVICE WARRANTY ONLY'
        $txtIn.Focus()
    }
}


# --------------------------------------------------------------------------
#  Wiring
# --------------------------------------------------------------------------
$tabWarranty.Add_Click({ Set-Mode 'Warranty' })
$tabParts.Add_Click({ Set-Mode 'Parts' })

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

# ---- parts section ----
$btnFind.Add_Click({ Invoke-PartLookup })

$btnCopyPart.Add_Click({
    $label = if (@($pgrid.SelectedRows).Count -eq 1 -or $script:PartMatches.Count -eq 1) { 'PART NO.' } else { 'PART NOS.' }
    Copy-ToClipboardSafe -Text (Get-PartNumbersText) -Label $label -Target $pstatus
})

$btnCopyPTable.Add_Click({ Copy-ToClipboardSafe -Text (Get-PartTableText) -Label 'TABLE' -Target $pstatus })

$btnPClear.Add_Click({
    $txtPSerial.Clear()
    $pgrid.Rows.Clear()
    $script:PartsList   = $null
    $script:PartMatches = @()
    Update-PartInfo
    foreach ($b in @($btnCopyPart, $btnCopyPTable)) { Set-BlockEnabled $b $false }
    Set-Status 'READY' $navy -Target $pstatus
    $txtPSerial.Focus()
})

# Changing the part re-filters the list already downloaded for that serial.
$cboPart.Add_SelectedIndexChanged({
    if (-not $script:PartsList -or $script:PartsList.Error) { return }
    $typed = @(ConvertTo-SerialList -Text $txtPSerial.Text)
    if ($typed.Count -ge 1 -and $typed[0] -eq $script:PartsList.Serial) { Update-PartGrid }
})

# Enter (or Ctrl+Enter) in the serial box runs the lookup.
$txtPSerial.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $_.SuppressKeyPress = $true
        Invoke-PartLookup
    }
})

$form.Add_Shown({
    foreach ($b in @($btnCopyDates, $btnCopyTable, $btnSaveCsv, $btnCopyPart, $btnCopyPTable)) { Set-BlockEnabled $b $false }
    Set-Mode 'Warranty'
})

[void]$form.ShowDialog()
$form.Dispose()
