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
    Save-Xml $parts[1] $doc

    $wbText = $wb.OuterXml
    if ($wbText -notmatch 'calcPr') {
        $calc = $wb.CreateElement('x', 'calcPr', $nsUri); $calc.SetAttribute('fullCalcOnLoad', '1')
        $sheets = $wb.GetElementsByTagName('sheets', $nsUri)[0]; [void]$wb.DocumentElement.InsertAfter($calc, $sheets)
        Save-Xml 'xl/workbook.xml' $wb
    }
}
finally { $zip.Dispose() }
Write-Host "Workbook updated from CSV: $($books[0].Name)"
