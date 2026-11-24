<#
.SYNOPSIS
  Copies the CSV contents into the Excel workbook, so the workbook never has to be edited in Excel.
.DESCRIPTION
  - main sheet (1st)    <- data/TISZA_merleg.csv
  - source sheet (3rd)  <- data/Forrasjegyzek.csv
  - summary sheet (2nd) <- cached COUNTIF results recomputed from TISZA_merleg.csv
  Cell styles are kept; new rows copy the style of the last existing row, new columns the
  style of the cell to their left. Rows beyond the CSV are removed. Run validate.ps1 afterwards.
.EXAMPLE
  pwsh -File scripts/sync_xlsx.ps1        # or: powershell -File scripts\sync_xlsx.ps1
#>
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

$root = Split-Path -Parent $PSScriptRoot
$data = Join-Path $root 'data'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$nsUri = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$relNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

function Read-Csv([string]$name) {
    $text = [IO.File]::ReadAllText((Join-Path $data $name), $utf8) -replace "`r`n", "`n"
    $count = (($text -split "`n", 2)[0] -split ',').Count
    $names = 0..($count - 1) | ForEach-Object { "c$_" }
    $rows = @()
    foreach ($r in @($text | ConvertFrom-Csv -Header $names)) { $rows += , @($names | ForEach-Object { if ($null -eq $r.$_) { '' } else { [string]$r.$_ } }) }
    , $rows
}
function Col-Letter([int]$j) { $s = ''; $n = $j + 1; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [string][char](65 + $m) + $s; $n = [int](($n - $m - 1) / 26) }; $s }

$books = @(Get-ChildItem $data -Filter '*.xlsx')
if ($books.Count -ne 1) { throw "data/: expected exactly one .xlsx, found $($books.Count)" }
$main = Read-Csv 'TISZA_merleg.csv'
$src = Read-Csv 'Forrasjegyzek.csv'

$zip = [IO.Compression.ZipFile]::Open($books[0].FullName, 'Update')
try {
    function Get-Xml([string]$entry) {
        $sr = New-Object IO.StreamReader($zip.GetEntry($entry).Open(), $utf8)
        try { $d = New-Object Xml.XmlDocument; $d.PreserveWhitespace = $true; $d.LoadXml($sr.ReadToEnd()); $d } finally { $sr.Dispose() }
    }
    function Save-Xml([string]$entry, [Xml.XmlDocument]$doc) {
        $sw = New-Object IO.StringWriter; $xs = New-Object Xml.XmlWriterSettings; $xs.NewLineHandling = 'None'
        $xw = [Xml.XmlWriter]::Create($sw, $xs); $doc.Save($xw); $xw.Dispose()
        $text = $sw.ToString() -replace '^<\?xml version="1.0" encoding="utf-16"\?>', '<?xml version="1.0" encoding="utf-8"?>'
        $s = $zip.GetEntry($entry).Open(); $s.SetLength(0); $b = $utf8.GetBytes($text); $s.Write($b, 0, $b.Length); $s.Dispose()
    }

    # sheet order -> part names
    $wb = Get-Xml 'xl/workbook.xml'
    $rels = Get-Xml 'xl/_rels/workbook.xml.rels'
    $parts = @(foreach ($s in $wb.GetElementsByTagName('sheet', $nsUri)) {
        $rid = $s.GetAttribute('id', $relNs)
        $t = ($rels.DocumentElement.ChildNodes | Where-Object { $_.LocalName -eq 'Relationship' -and $_.GetAttribute('Id') -eq $rid }).GetAttribute('Target').TrimStart('/')
        if ($t -notlike 'xl/*') { 'xl/' + $t } else { $t }
    })

    function Sync-Sheet([string]$part, $rows) {
        $doc = Get-Xml $part
        $ns = New-Object Xml.XmlNamespaceManager $doc.NameTable; $ns.AddNamespace('x', $nsUri)
        $sheetData = $doc.SelectSingleNode('//x:sheetData', $ns)
        $xrows = @($sheetData.SelectNodes('x:row', $ns))
        if ($xrows.Count -lt 2) { throw "${part}: needs a header row and at least one data row as style template" }
        # extra rows go, missing rows are cloned from the last one
        for ($i = $xrows.Count - 1; $i -ge $rows.Count; $i--) { [void]$sheetData.RemoveChild($xrows[$i]) }
        for ($i = $xrows.Count; $i -lt $rows.Count; $i++) { [void]$sheetData.AppendChild($xrows[-1].CloneNode($true)) }
        $xrows = @($sheetData.SelectNodes('x:row', $ns))
        for ($i = 0; $i -lt $rows.Count; $i++) {
            $xr = $xrows[$i]; $n = $i + 1; $xr.SetAttribute('r', [string]$n)
            $cells = @{}
            foreach ($c in @($xr.SelectNodes('x:c', $ns))) {
                $letters = $c.GetAttribute('r') -replace '\d', ''
                $c.SetAttribute('r', "$letters$n"); $cells[$letters] = $c
            }
            for ($j = 0; $j -lt $rows[$i].Count; $j++) {
                $L = Col-Letter $j; $val = $rows[$i][$j]
                if (-not $cells.ContainsKey($L)) {
                    $c = $doc.CreateElement('x', 'c', $nsUri); $c.SetAttribute('r', "$L$n")
                    if ($j -gt 0 -and $cells.ContainsKey((Col-Letter ($j - 1)))) { $c.SetAttribute('s', $cells[(Col-Letter ($j - 1))].GetAttribute('s')) }
                    [void]$xr.AppendChild($c); $cells[$L] = $c
                }
                $c = $cells[$L]
                foreach ($child in @($c.ChildNodes)) { [void]$c.RemoveChild($child) }
                if ($val -eq '') { $c.SetAttribute('t', 'n'); continue }
                if ($i -gt 0 -and $j -eq 0 -and $val -match '^\d+$') { $c.SetAttribute('t', 'n') }
                else { $c.SetAttribute('t', 'str') }
                $v = $doc.CreateElement('x', 'v', $nsUri); $v.InnerText = $val; [void]$c.AppendChild($v)
            }
        }
        # width for columns without a definition
        $cols = $doc.SelectSingleNode('//x:cols', $ns)
        if ($cols) {
            $defined = @($cols.SelectNodes('x:col', $ns) | ForEach-Object { [int]$_.GetAttribute('max') } | Measure-Object -Maximum).Maximum
            for ($k = [int]$defined + 1; $k -le $rows[0].Count; $k++) {
                $col = $doc.CreateElement('x', 'col', $nsUri)
                foreach ($a in @(@('min', "$k"), @('max', "$k"), @('width', '16'), @('customWidth', '1'))) { $col.SetAttribute($a[0], $a[1]) }
                [void]$cols.AppendChild($col)
            }
        }
        Save-Xml $part $doc
        "{0}: {1} rows x {2} columns" -f $part, $rows.Count, $rows[0].Count
    }

    Sync-Sheet $parts[0] $main
    Sync-Sheet $parts[2] $src

    function U([string]$s) { [regex]::Unescape($s) }

    # --- main sheet: status colours taken from the summary sheet, A:B frozen -------------
    # Each status label in column A of the summary sheet carries the fill to use; every row of
    # the main sheet gets that fill (base cell styles: A2 for the row, E2 for the status column).
    $sty = Get-Xml 'xl/styles.xml'
    $sns = New-Object Xml.XmlNamespaceManager $sty.NameTable; $sns.AddNamespace('x', $nsUri)
    $cellXfs = $sty.SelectSingleNode('//x:cellXfs', $sns)
    function Get-XfWithFill([int]$base, [string]$fillId) {
        $xfs = @($cellXfs.SelectNodes('x:xf', $sns))
        $want = $xfs[$base].CloneNode($true); $want.SetAttribute('fillId', $fillId); $want.SetAttribute('applyFill', '1')
        for ($k = 0; $k -lt $xfs.Count; $k++) { if ($xfs[$k].OuterXml -eq $want.OuterXml) { return $k } }
        [void]$cellXfs.AppendChild($want); $cellXfs.SetAttribute('count', [string]($xfs.Count + 1))
        return $xfs.Count
    }
    function Get-FillOf([int]$xf) { @($cellXfs.SelectNodes('x:xf', $sns))[$xf].GetAttribute('fillId') }

    $sumDoc = Get-Xml $parts[1]
    $sumNs = New-Object Xml.XmlNamespaceManager $sumDoc.NameTable; $sumNs.AddNamespace('x', $nsUri)
    $statusCol = [array]::IndexOf($main[0], ($main[0] | Where-Object { $_ -like 'St*tusz' } | Select-Object -First 1))
    # statuses: those used in the main table plus those listed in the summary's COUNTIF table
    # (a status with 0 items still needs its legend colour)
    $countLabels = @($sumDoc.SelectNodes('//x:c[x:f]', $sumNs) | Where-Object { $_.SelectSingleNode('x:f', $sumNs).InnerText -match '^COUNTIF\(' } |
        ForEach-Object { $rn = $_.GetAttribute('r') -replace '[A-Z]', ''; $sumDoc.SelectSingleNode("//x:c[@r='A$rn']/x:v", $sumNs).InnerText })
    $labels = @(@($main | Select-Object -Skip 1 | ForEach-Object { $_[$statusCol] }) + $countLabels | Sort-Object -Unique)
    $statusFill = @{}
    foreach ($c in @($sumDoc.SelectNodes('//x:sheetData/x:row/x:c', $sumNs))) {
        if ($c.GetAttribute('r') -notmatch '^A\d+$') { continue }
        $v = $c.SelectSingleNode('x:v', $sumNs); if (-not $v) { continue }
        if ($labels -contains $v.InnerText -and $c.GetAttribute('s') -ne '') { $statusFill[$v.InnerText] = Get-FillOf ([int]$c.GetAttribute('s')) }
    }
    foreach ($l in $labels) { if (-not $statusFill.ContainsKey($l)) { throw "summary sheet has no coloured cell for status '$l'" } }

    $doc = Get-Xml $parts[0]
    $ns = New-Object Xml.XmlNamespaceManager $doc.NameTable; $ns.AddNamespace('x', $nsUri)
    $statusLetter = Col-Letter $statusCol
    $baseRow = [int]$doc.SelectSingleNode("//x:c[@r='A2']", $ns).GetAttribute('s')
    $baseStatus = [int]$doc.SelectSingleNode("//x:c[@r='${statusLetter}2']", $ns).GetAttribute('s')
    $rowsXml = @($doc.SelectNodes('//x:sheetData/x:row', $ns))
    for ($i = 1; $i -lt $rowsXml.Count; $i++) {
        $fid = $statusFill[$main[$i][$statusCol]]
        $rowXf = Get-XfWithFill $baseRow $fid; $statXf = Get-XfWithFill $baseStatus $fid
        foreach ($c in @($rowsXml[$i].SelectNodes('x:c', $ns))) {
            $c.SetAttribute('s', [string]$(if (($c.GetAttribute('r') -replace '\d', '') -eq $statusLetter) { $statXf } else { $rowXf }))
        }
    }
    # freeze the header row and columns A:B
    $view = $doc.SelectSingleNode('//x:sheetViews/x:sheetView', $ns)
    foreach ($n in @($view.ChildNodes)) { [void]$view.RemoveChild($n) }
    $pane = $doc.CreateElement('x', 'pane', $nsUri)
    foreach ($a in @(@('xSplit', '2'), @('ySplit', '1'), @('topLeftCell', 'C2'), @('activePane', 'bottomRight'), @('state', 'frozen'))) { $pane.SetAttribute($a[0], $a[1]) }
    [void]$view.AppendChild($pane)
    $sel = $doc.CreateElement('x', 'selection', $nsUri)
    foreach ($a in @(@('pane', 'bottomRight'), @('activeCell', 'C2'), @('sqref', 'C2'))) { $sel.SetAttribute($a[0], $a[1]) }
    [void]$view.AppendChild($sel)
    Save-Xml $parts[0] $doc

    # legend on the summary sheet: "<status> - <meaning>" lines get the colour of their status
    foreach ($c in @($sumDoc.SelectNodes('//x:sheetData/x:row/x:c', $sumNs))) {
        if ($c.GetAttribute('r') -notmatch '^A\d+$') { continue }
        $v = $c.SelectSingleNode('x:v', $sumNs); if (-not $v) { continue }
        foreach ($l in $labels) {
            if ($v.InnerText.StartsWith($l + ' ' + [char]0x2013)) { $c.SetAttribute('s', [string](Get-XfWithFill $baseRow $statusFill[$l])) }
        }
    }
    Save-Xml $parts[1] $sumDoc
    Save-Xml 'xl/styles.xml' $sty

    # --- project sheet: rows for item count, deadlines and data flow ----------------------
    if ($parts.Count -ge 4) {
        $doc = Get-Xml $parts[3]
        $ns = New-Object Xml.XmlNamespaceManager $doc.NameTable; $ns.AddNamespace('x', $nsUri)
        $sd = $doc.SelectSingleNode('//x:sheetData', $ns)
        $text = { param($row, $col) $n = $row.SelectSingleNode("x:c[starts-with(@r,'$col')]/x:v", $ns); if ($n) { $n.InnerText } else { '' } }
        $rowsXml = @($sd.SelectNodes('x:row', $ns))
        $note = $rowsXml | Where-Object { (& $text $_ 'A').StartsWith((U 'Megjegyz\u00e9s')) } | Select-Object -First 1
        $dataRows = @($rowsXml | Where-Object { [int]$_.GetAttribute('r') -ge 4 -and $_ -ne $note -and (& $text $_ 'A') -ne '' })
        $template = $dataRows | Where-Object { $_.SelectSingleNode("x:c[starts-with(@r,'D')]", $ns) } | Select-Object -First 1
        $wanted = @(
            @((U 'T\u00e9telek sz\u00e1ma'), [string]($main.Count - 1), (U 'A st\u00e1tuszok szerinti bont\u00e1s az \u00d6sszefoglal\u00f3 munkalapon.'), ''),
            @((U 'Hat\u00e1rid\u0151k'), (U 'V\u00e1llalt hat\u00e1rid\u0151 oszlop'), (U 'Csak a programban szerepl\u0151 hat\u00e1rid\u0151: els\u0151 int\u00e9zked\u00e9sek, azonnal, konkr\u00e9t \u00e9v vagy nincs megadva. \u274c csak a hat\u00e1rid\u0151 ut\u00e1n adhat\u00f3.'), 'https://github.com/cooldance84/tisza-kormany-merleg#hat%C3%A1rid%C5%91k-%C3%A9s-a-t%C3%A9telek-%C3%A9letciklusa'),
            @((U 'Adatok'), (U 'CSV \u2192 Excel'), (U 'A munkaf\u00fczet a data/*.csv f\u00e1jlokb\u00f3l k\u00e9sz\u00fcl (scripts/sync_xlsx.ps1); szerkeszteni a CSV-ben kell.'), 'https://github.com/cooldance84/tisza-kormany-merleg/tree/main/data')
        )
        foreach ($w in $wanted) {
            $row = $dataRows | Where-Object { (& $text $_ 'A') -eq $w[0] } | Select-Object -First 1
            if (-not $row) { $row = $template.CloneNode($true); [void]$sd.InsertBefore($row, $note); $dataRows += $row }
            foreach ($j in 0..3) {
                $L = Col-Letter $j
                $c = $row.SelectSingleNode("x:c[starts-with(@r,'$L')]", $ns); if (-not $c) { continue }
                foreach ($child in @($c.ChildNodes)) { [void]$c.RemoveChild($child) }
                if ($w[$j] -eq '') { $c.SetAttribute('t', 'n'); continue }
                $c.SetAttribute('t', 'str'); $v = $doc.CreateElement('x', 'v', $nsUri); $v.InnerText = $w[$j]; [void]$c.AppendChild($v)
            }
        }
        # renumber: data rows from 4, one blank row, then the note; keep the note's merge in step
        $oldNote = $note.GetAttribute('r'); $n = 4
        foreach ($row in @($sd.SelectNodes('x:row', $ns))) {
            $r = [int]$row.GetAttribute('r')
            if ($row -eq $note -or $r -lt 4) { continue }
            $row.SetAttribute('r', [string]$n)
            foreach ($c in @($row.SelectNodes('x:c', $ns))) { $c.SetAttribute('r', ($c.GetAttribute('r') -replace '\d+$', '') + $n) }
            $n++
        }
        $newNote = $n + 1
        $note.SetAttribute('r', [string]$newNote)
        foreach ($c in @($note.SelectNodes('x:c', $ns))) { $c.SetAttribute('r', ($c.GetAttribute('r') -replace '\d+$', '') + $newNote) }
        [void]$sd.RemoveChild($note); [void]$sd.AppendChild($note)
        foreach ($mc in @($doc.SelectNodes('//x:mergeCell', $ns))) {
            if ($mc.GetAttribute('ref') -eq "A${oldNote}:D${oldNote}") { $mc.SetAttribute('ref', "A${newNote}:D${newNote}") }
        }
        $dim = $doc.SelectSingleNode('//x:dimension', $ns); if ($dim) { $dim.SetAttribute('ref', "A1:D$newNote") }
        Save-Xml $parts[3] $doc
    }

    # summary sheet: cached COUNTIF values (Excel recalculates on open, other readers use the cache)
    $doc = Get-Xml $parts[1]
    $ns = New-Object Xml.XmlNamespaceManager $doc.NameTable; $ns.AddNamespace('x', $nsUri)
    $statusCol = [array]::IndexOf($main[0], ($main[0] | Where-Object { $_ -like 'St*tusz' } | Select-Object -First 1))
    foreach ($fc in @($doc.SelectNodes('//x:c[x:f]', $ns))) {
        if ($fc.SelectSingleNode('x:f', $ns).InnerText -notmatch '^COUNTIF\(') { continue }
        $rn = $fc.GetAttribute('r') -replace '[A-Z]', ''
        $label = $doc.SelectSingleNode("//x:c[@r='A$rn']", $ns).InnerText
        $count = @($main | Select-Object -Skip 1 | Where-Object { $_[$statusCol] -eq $label }).Count
        $v = $fc.SelectSingleNode('x:v', $ns); if (-not $v) { $v = $doc.CreateElement('x', 'v', $nsUri); [void]$fc.AppendChild($v) }
        $v.InnerText = [string]$count
    }

    # "last updated" stamps: the latest last-checked date of the main table, as YYYY.MM.DD.
    # (\u escapes keep this file ASCII-only for Windows PowerShell 5.1)
    $checkedCol = [array]::IndexOf($main[0], ($main[0] | Where-Object { $_ -like 'Utols*ellen*' } | Select-Object -First 1))
    $stamp = (@($main | Select-Object -Skip 1 | ForEach-Object { $_[$checkedCol] }) | Sort-Object | Select-Object -Last 1).Replace('-', '.') + '.'
    $titleRe = [regex]::Unescape('friss\u00edtve: ') + '\d{4}\.\d{2}\.\d{2}\.'
    foreach ($v in @($doc.SelectNodes('//x:c/x:v', $ns))) {
        if ($v.InnerText -match $titleRe) { $v.InnerText = $v.InnerText -replace $titleRe, ([regex]::Unescape('friss\u00edtve: ') + $stamp) }
    }
    Save-Xml $parts[1] $doc

    # project sheet (4th): "Utolso frissites" row -> date and a pointer to the changelog
    if ($parts.Count -ge 4) {
        $doc = Get-Xml $parts[3]
        $ns = New-Object Xml.XmlNamespaceManager $doc.NameTable; $ns.AddNamespace('x', $nsUri)
        foreach ($row in @($doc.SelectNodes('//x:sheetData/x:row', $ns))) {
            $rn = $row.GetAttribute('r')
            $a = $row.SelectSingleNode("x:c[@r='A$rn']/x:v", $ns)
            if (-not $a -or $a.InnerText -ne [regex]::Unescape('Utols\u00f3 friss\u00edt\u00e9s')) { continue }
            $b = $row.SelectSingleNode("x:c[@r='B$rn']/x:v", $ns); if ($b) { $b.InnerText = $stamp }
            $c = $row.SelectSingleNode("x:c[@r='C$rn']/x:v", $ns)
            if ($c) { $c.InnerText = [regex]::Unescape('A v\u00e1ltoz\u00e1sok r\u00e9szletei: CHANGELOG.md') }
        }
        Save-Xml $parts[3] $doc
    }

    $wbText = $wb.OuterXml
    if ($wbText -notmatch 'calcPr') {
        $calc = $wb.CreateElement('x', 'calcPr', $nsUri); $calc.SetAttribute('fullCalcOnLoad', '1')
        $sheets = $wb.GetElementsByTagName('sheets', $nsUri)[0]; [void]$wb.DocumentElement.InsertAfter($calc, $sheets)
        Save-Xml 'xl/workbook.xml' $wb
    }
}
finally { $zip.Dispose() }
Write-Host "Workbook updated from CSV: $($books[0].Name)"
