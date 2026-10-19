<#
.SYNOPSIS
  Consistency checks for the data/ folder: CSV structure, status counts,
  item counts quoted in the READMEs, and CSV <-> xlsx sync.
.EXAMPLE
  pwsh -File scripts/validate.ps1        # or: powershell -File scripts\validate.ps1
#>
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

$root = Split-Path -Parent $PSScriptRoot
$data = Join-Path $root 'data'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$errors = New-Object System.Collections.Generic.List[string]
function Fail([string]$msg) { $errors.Add($msg) }

# Reads a CSV by column position: returns header names and rows as string arrays.
function Read-CsvFile([string]$name, [int]$columns) {
    $path = Join-Path $data $name
    $text = [IO.File]::ReadAllText($path, $utf8) -replace "`r`n", "`n"
    $headerLine = ($text -split "`n", 2)[0]
    $header = $headerLine -split ','
    if ($header.Count -ne $columns) { Fail "${name}: header has $($header.Count) columns, expected $columns" }
    $names = 0..($columns - 1) | ForEach-Object { "c$_" }
    $rows = @($text | ConvertFrom-Csv -Header $names | Select-Object -Skip 1 | ForEach-Object {
        $r = $_; , @($names | ForEach-Object { if ($null -eq $r.$_) { '' } else { [string]$r.$_ } })
    })
    [pscustomobject]@{ Name = $name; Header = $header; Rows = $rows }
}

# Reads an xlsx sheet as rows of cell strings (inline, shared and plain values).
function Read-XlsxSheets([string]$path) {
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $read = {
            param($entryName)
            $e = $zip.GetEntry($entryName); if (-not $e) { return $null }
            $sr = New-Object IO.StreamReader($e.Open(), $utf8); try { [xml]$sr.ReadToEnd() } finally { $sr.Dispose() }
        }
        $nsUri = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        $wb = & $read 'xl/workbook.xml'
        $ns = New-Object Xml.XmlNamespaceManager $wb.NameTable; $ns.AddNamespace('x', $nsUri)
        $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $rels = & $read 'xl/_rels/workbook.xml.rels'
        $shared = @()
        $sst = & $read 'xl/sharedStrings.xml'
        if ($sst) {
            $sns = New-Object Xml.XmlNamespaceManager $sst.NameTable; $sns.AddNamespace('x', $nsUri)
            $shared = @($sst.SelectNodes('//x:si', $sns) | ForEach-Object { ($_.SelectNodes('.//x:t', $sns) | ForEach-Object { $_.InnerText }) -join '' })
        }
        $sheets = [ordered]@{}
        foreach ($s in $wb.SelectNodes('//x:sheets/x:sheet', $ns)) {
            $rid = $s.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $target = ($rels.Relationships.Relationship | Where-Object { $_.Id -eq $rid }).Target.TrimStart('/')
            if ($target -notlike 'xl/*') { $target = 'xl/' + $target }
            $doc = & $read $target
            $dns = New-Object Xml.XmlNamespaceManager $doc.NameTable; $dns.AddNamespace('x', $nsUri)
            $rows = New-Object System.Collections.Generic.List[object]
            foreach ($row in $doc.SelectNodes('//x:sheetData/x:row', $dns)) {
                $cells = @{}; $max = -1
                foreach ($c in $row.SelectNodes('x:c', $dns)) {
                    $letters = ($c.GetAttribute('r') -replace '\d', '')
                    $col = 0; foreach ($ch in $letters.ToCharArray()) { $col = $col * 26 + ([int]$ch - 64) }; $col--
                    $t = $c.GetAttribute('t'); $v = $c.SelectSingleNode('x:v', $dns)
                    $val = switch ($t) {
                        'inlineStr' { ($c.SelectNodes('x:is//x:t', $dns) | ForEach-Object { $_.InnerText }) -join '' }
                        's' { $shared[[int]$v.InnerText] }
                        default { if ($v) { $v.InnerText } else { '' } }
                    }
                    $cells[$col] = ($val -replace "`r`n", "`n"); if ($col -gt $max) { $max = $col }
                }
                $arr = @(0..([Math]::Max($max, 0)) | ForEach-Object { if ($cells.ContainsKey($_)) { $cells[$_] } else { '' } })
                if (($arr -join '') -ne '') { $rows.Add($arr) }
            }
            $formulas = @{}
            foreach ($c in $doc.SelectNodes('//x:c[x:f]', $dns)) { $formulas[$c.GetAttribute('r')] = $c.SelectSingleNode('x:f', $dns).InnerText }
            $merges = @($doc.SelectNodes('//x:mergeCells/x:mergeCell', $dns) | ForEach-Object { $_.GetAttribute('ref') })
            $sheets[$s.GetAttribute('name')] = [pscustomobject]@{ Rows = $rows.ToArray(); Formulas = $formulas; Merges = $merges }
        }
        $sheets
    } finally { $zip.Dispose() }
}

function Compare-Table($label, $csv, $sheetRows, [int]$columns) {
    $csvRows = @(, $csv.Header) + $csv.Rows
    if ($sheetRows.Count -ne $csvRows.Count) { Fail "${label}: xlsx has $($sheetRows.Count) non-empty rows, CSV has $($csvRows.Count)" }
    $n = [Math]::Min($sheetRows.Count, $csvRows.Count)
    for ($i = 0; $i -lt $n; $i++) {
        for ($j = 0; $j -lt $columns; $j++) {
            $a = if ($j -lt $sheetRows[$i].Count) { $sheetRows[$i][$j] } else { '' }
            $b = if ($j -lt $csvRows[$i].Count) { $csvRows[$i][$j] } else { '' }
            if ($a -ne $b) { Fail ("{0}: row {1} column {2} differs`n    xlsx: {3}`n    csv:  {4}" -f $label, ($i + 1), ($j + 1), $a, $b) }
        }
    }
}

# --- main table ---------------------------------------------------------------
$main = Read-CsvFile 'TISZA_merleg.csv' 9
if ($main.Header[0] -ne 'ID') { Fail "TISZA_merleg.csv: first column must be 'ID' (a '#' header is dropped as a comment by PowerShell Import-Csv)" }
$urlLine = '^([^:]+: )?https://\S+$'
for ($i = 0; $i -lt $main.Rows.Count; $i++) {
    $r = $main.Rows[$i]; $id = $r[0]
    if ($id -ne [string]($i + 1)) { Fail "TISZA_merleg.csv: row $($i + 1) has ID '$id', expected $($i + 1) (IDs must be 1..N without gaps)" }
    foreach ($j in 1..4) { if ($r[$j].Trim() -eq '') { Fail "TISZA_merleg.csv: ID $id column '$($main.Header[$j])' is empty" } }
    $lines = @($r[6] -split "`n" | Where-Object { $_.Trim() -ne '' })
    if ($lines.Count -eq 0) { Fail "TISZA_merleg.csv: ID $id has no source" }
    foreach ($l in $lines) { if ($l -notmatch $urlLine) { Fail "TISZA_merleg.csv: ID $id source line is not '[Label: ]https://...': $l" } }
    $checked = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($r[8], 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$checked)) {
        Fail "TISZA_merleg.csv: ID $id last-checked date '$($r[8])' is not YYYY-MM-DD"
    } elseif ($checked -gt [datetime]::Today.AddDays(1)) { Fail "TISZA_merleg.csv: ID $id last-checked date $($r[8]) is in the future" }
}
$itemCount = $main.Rows.Count

# --- status summary -----------------------------------------------------------
$sum = Read-CsvFile 'Osszefoglalo.csv' 3
$known = @{}
foreach ($r in $sum.Rows) { $known[$r[0]] = [int]$r[2] }
$actual = @{}
foreach ($r in $main.Rows) {
    if (-not $known.ContainsKey($r[4])) { Fail "TISZA_merleg.csv: ID $($r[0]) has unknown status '$($r[4])'" }
    $actual[$r[4]] = 1 + $(if ($actual.ContainsKey($r[4])) { $actual[$r[4]] } else { 0 })
}
foreach ($k in $known.Keys) {
    $a = if ($actual.ContainsKey($k)) { $actual[$k] } else { 0 }
    if ($a -ne $known[$k]) { Fail "Osszefoglalo.csv: '$k' says $($known[$k]), TISZA_merleg.csv has $a" }
}

# --- source index ---------------------------------------------------------------
$src = Read-CsvFile 'Forrasjegyzek.csv' 5
foreach ($r in $src.Rows) {
    foreach ($l in @($r[4] -split "`n" | Where-Object { $_.Trim() -ne '' })) {
        if ($l -notmatch '^https://\S+$') { Fail "Forrasjegyzek.csv: '$($r[0])' Link is not an https URL: $l" }
    }
    if ($r[3] -match 'https?://') { Fail "Forrasjegyzek.csv: '$($r[0])' has a URL in the note column; put it in Link" }
}

# --- item counts quoted in docs ---------------------------------------------------
foreach ($doc in 'README.md', 'data/README.md') {
    $text = [IO.File]::ReadAllText((Join-Path $root $doc), $utf8)
    # \u00e9 = e-acute; keeps this file ASCII-only so Windows PowerShell 5.1 reads it correctly without a BOM
    foreach ($m in [regex]::Matches($text, '(\d+)\s+t\u00e9tel')) {
        $n = [int]$m.Groups[1].Value
        if ($n -ne $itemCount) { Fail "${doc}: mentions '$($m.Value)', but TISZA_merleg.csv has $itemCount items" }
    }
}

# --- xlsx <-> CSV -----------------------------------------------------------------
$books = @(Get-ChildItem $data -Filter '*.xlsx')
if ($books.Count -ne 1) { Fail "data/: expected exactly one .xlsx, found $($books.Count)" }
else {
    # A Microsoft sensitivity label (MSIP) carries the labelling organisation's tenant and label ids;
    # an encrypting label turns the file into a non-zip container. Neither belongs in this public repo.
    try { $zip = [IO.Compression.ZipFile]::OpenRead($books[0].FullName) }
    catch { Fail "$($books[0].Name): not a readable xlsx (zip) file - encrypted by a sensitivity label?"; $zip = $null }
    if ($zip) {
        try {
            foreach ($e in $zip.Entries) {
                if ($e.FullName -like 'docMetadata/*') { Fail "$($books[0].Name): contains $($e.FullName) (sensitivity label metadata); remove the label" ; continue }
                $sr = New-Object IO.StreamReader($e.Open(), $utf8)
                try { if ($sr.ReadToEnd() -match 'MSIP_Label_') { Fail "$($books[0].Name): $($e.FullName) contains MSIP_Label_ properties (sensitivity label); remove the label" } }
                finally { $sr.Dispose() }
            }
        } finally { $zip.Dispose() }
    }
    $sheets = if ($zip) { Read-XlsxSheets $books[0].FullName } else { [ordered]@{} }
    $names = @($sheets.Keys)
    if (-not $zip) { }
    elseif ($names.Count -lt 3) { Fail "$($books[0].Name): expected at least 3 sheets, found $($names.Count)" }
    else {
        Compare-Table "$($books[0].Name) [$($names[0])]" $main $sheets[$names[0]].Rows 9
        Compare-Table "$($books[0].Name) [$($names[2])]" $src $sheets[$names[2]].Rows 5
        # Overlapping merged ranges make Excel "repair" the file on open and drop them.
        foreach ($name in $names) {
            $boxes = @($sheets[$name].Merges | ForEach-Object {
                $m = [regex]::Match($_, '^([A-Z]+)(\d+):([A-Z]+)(\d+)$')
                $col = { param($l) $v = 0; foreach ($ch in $l.ToCharArray()) { $v = $v * 26 + ([int]$ch - 64) }; $v }
                , @($_, (& $col $m.Groups[1].Value), [int]$m.Groups[2].Value, (& $col $m.Groups[3].Value), [int]$m.Groups[4].Value)
            })
            for ($i = 0; $i -lt $boxes.Count; $i++) {
                for ($j = $i + 1; $j -lt $boxes.Count; $j++) {
                    $a = $boxes[$i]; $b = $boxes[$j]
                    if ($a[1] -le $b[3] -and $b[1] -le $a[3] -and $a[2] -le $b[4] -and $b[2] -le $a[4]) {
                        Fail "$($books[0].Name) [$name]: merged ranges $($a[0]) and $($b[0]) overlap (Excel repairs the file on open)"
                    }
                }
            }
        }
        # Summary sheet layout differs from the CSV; check its COUNTIF ranges are not capped at a fixed row.
        foreach ($f in $sheets[$names[1]].Formulas.GetEnumerator()) {
            if ($f.Value -match 'COUNTIF\(.*\$E\$\d+:\$E\$(\d+)' -and [int]$Matches[1] -lt 1000) {
                Fail "$($books[0].Name) [$($names[1])]!$($f.Key): range ends at row $($Matches[1]); use `$E:`$E"
            }
        }
    }
}

if ($errors.Count) {
    $errors | ForEach-Object { Write-Host "ERROR: $_" -ForegroundColor Red }
    Write-Host "$($errors.Count) problem(s) found."
    exit 1
}
Write-Host "OK: $itemCount items, $($known.Count) statuses, $($src.Rows.Count) sources, CSV and xlsx in sync."
