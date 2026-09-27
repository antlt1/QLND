# Thu vien dung chung: doc XLSX truc tiep (khong can Excel/Node/Python)
Add-Type -AssemblyName System.IO.Compression.FileSystem

function U([string]$h) {
    $sb = ""
    foreach ($p in $h.Split(' ')) { $sb += [char][Convert]::ToInt32($p, 16) }
    return $sb
}
# chu ho / thanh vien (nhung bang \uXXXX de file script luon la ASCII)
$CHU_HO  = U "63 68 1EE7 20 68 1ED9"
$TV      = U "74 68 E0 6E 68 20 76 69 EA 6E"
$KHONG_CO= U "6B 68 F4 6E 67 20 63 F3"
$REL     = @($CHU_HO, $TV)

function Get-EntryText($zip, $name) {
    $e = $zip.Entries | Where-Object { $_.FullName -eq $name }
    if ($null -eq $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    $t = $sr.ReadToEnd(); $sr.Close(); return $t
}

# Doc 1 sheet cua .xlsx -> mang object {Row, C} voi C = hashtable cot(1-based) -> chuoi
function Read-Sheet($Path, $SheetIndex) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $sst = @()
        $sx = Get-EntryText $zip "xl/sharedStrings.xml"
        if ($sx) {
            $d = New-Object System.Xml.XmlDocument; $d.LoadXml($sx)
            foreach ($si in $d.DocumentElement.ChildNodes) {
                $s = ""
                foreach ($t in $si.SelectNodes(".//*[local-name()='t']")) { $s += $t.InnerText }
                $sst += $s
            }
        }
        $dw = New-Object System.Xml.XmlDocument; $dw.LoadXml((Get-EntryText $zip "xl/workbook.xml"))
        $dr = New-Object System.Xml.XmlDocument; $dr.LoadXml((Get-EntryText $zip "xl/_rels/workbook.xml.rels"))
        $rm = @{}
        foreach ($rel in $dr.DocumentElement.ChildNodes) { $rm[$rel.Id] = $rel.Target }
        $sh = @($dw.DocumentElement.SelectNodes("//*[local-name()='sheet']"))[$SheetIndex]
        $rid = $sh.GetAttribute("id"); if (-not $rid) { $rid = $sh.GetAttribute("r:id") }
        $tg = $rm[$rid]; if ($tg -notlike "xl/*") { $tg = "xl/" + $tg.TrimStart('/') }
        $ds = New-Object System.Xml.XmlDocument; $ds.LoadXml((Get-EntryText $zip $tg))
        $res = @()
        foreach ($row in $ds.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
            $rn = [int]$row.GetAttribute("r")
            $cells = @{}
            foreach ($c in $row.ChildNodes) {
                if ($c.LocalName -ne "c") { continue }
                $ref = $c.GetAttribute("r"); $t = $c.GetAttribute("t")
                $vN = $c.SelectSingleNode("*[local-name()='v']")
                $iN = $c.SelectSingleNode("*[local-name()='is']")
                $val = ""
                if ($t -eq "s" -and $vN) { $val = $sst[[int]$vN.InnerText] }
                elseif ($t -eq "inlineStr" -and $iN) {
                    $x = ""; foreach ($tn in $iN.SelectNodes(".//*[local-name()='t']")) { $x += $tn.InnerText }; $val = $x
                }
                elseif ($vN) { $val = $vN.InnerText }
                $cl = ($ref -replace '[0-9]', ''); $ci = 0
                foreach ($ch in $cl.ToCharArray()) { $ci = $ci * 26 + ([int]$ch - 64) }
                $cells[$ci] = $val
            }
            $res += [pscustomobject]@{ Row = $rn; C = $cells }
        }
        return , $res
    } finally { $zip.Dispose() }
}

function Col($r, $i) { if ($r.C.ContainsKey($i)) { return $r.C[$i].Trim() } else { return "" } }

function Ser2Date($s) {
    $n = 0
    if ([double]::TryParse($s, [ref]$n) -and $n -gt 1000 -and $n -lt 80000) {
        return ([DateTime]::FromOADate($n)).ToString("yyyy-MM-dd")
    }
    return $null
}
function Txt2Date($s) {
    $s = $s.Trim()
    if ($s -match '^(\d{1,2})/(\d{1,2})/(\d{4})$') { return ("{0}-{1}-{2}" -f $Matches[3], $Matches[2].PadLeft(2,'0'), $Matches[1].PadLeft(2,'0')) }
    if ($s -match '^(\d{1,2})/(\d{1,2})/(\d{2})$') {
        $y = [int]$Matches[3]; if ($y -lt 30) { $y += 2000 } else { $y += 1900 }
        return ("{0}-{1}-{2}" -f $y, $Matches[2].PadLeft(2,'0'), $Matches[1].PadLeft(2,'0'))
    }
    if ($s -match '^(\d{4})-(\d{1,2})-(\d{1,2})$') { return ("{0}-{1}-{2}" -f $Matches[1], $Matches[2].PadLeft(2,'0'), $Matches[3].PadLeft(2,'0')) }
    return $null
}
function NormName($s) { return (($s -replace '\s+', ' ').Trim()).ToLower() }
function Save-Utf8($path, $text) {
    [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($true)))
}
