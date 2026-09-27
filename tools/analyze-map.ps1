param(
    [string]$SrcDir = "C:\Users\Administrator\Downloads\An\QLND\source\fileexcel",
    [string]$OutDir = "C:\Users\Administrator\Downloads\An\QLND\output"
)
$ErrorActionPreference = "Stop"
. "$PSScriptRoot\xlsx-lib.ps1"

$sb = New-Object System.Text.StringBuilder
function L($s) { [void]$sb.AppendLine($s) }

$to1File = $null; $traFile = $null
foreach ($f in (Get-ChildItem -LiteralPath $SrcDir -Filter "*.xlsx")) {
    if ($f.Name -like "* 1.xlsx" -and $f.Length -gt 100000) { $to1File = $f.FullName }
    if ($f.Name -like "*tra.xlsx" -and $f.Length -lt 140000) { $traFile = $f.FullName }
}

# ---- TOT 1: doc CCCD + STT ho hien tai (forward-fill) ----
$t1 = @()
$curA = ""; $curHo = 0
foreach ($r in (Read-Sheet $to1File 0)) {
    if ($r.Row -lt 3) { continue }
    $ten = Col $r 3
    if ($ten -eq "") { continue }
    $a = Col $r 1
    if ($a -ne "") { $curA = $a; if ($a -match '^\d+$') { $curHo = [int]$a } }
    $c = Col $r 6
    $cccd = $c; if ($c -notmatch '^\d{12}$') { $cccd = "" }
    $t1 += [pscustomobject]@{ row = $r.Row; ten = $ten; aRaw = $curA; ho = $curHo; cccd = $cccd }
}
L "To 1: $($t1.Count) nguoi"

# ---- file tra ----
$tra = @()
$curHh = 0
foreach ($r in (Read-Sheet $traFile 0)) {
    if ($r.Row -lt 2) { continue }
    $ten = Col $r 2
    if ($ten -eq "") { continue }
    if ((Col $r 7) -eq $CHU_HO) { $curHh++ }
    $cc = Col $r 5
    $tra += [pscustomobject]@{ stt = (Col $r 1); ten = $ten; cccd = $(if ($cc -match '^\d{12}$') { $cc } else { "" }); hoTra = $curHh }
}
L "File tra: $($tra.Count) nguoi / $curHh ho"

$traByCccd = @{}; $traByName = @{}
foreach ($p in $tra) {
    if ($p.cccd -ne "") { if (-not $traByCccd.ContainsKey($p.cccd)) { $traByCccd[$p.cccd] = $p } }
    $nk = NormName $p.ten
    if ($nk -ne "") { if (-not $traByName.ContainsKey($nk)) { $traByName[$nk] = $p } }
}

# ---- vung bi chon dong: r149..r325 ----
$shift = @($t1 | Where-Object { $_.row -ge 149 -and $_.row -le 325 })
L "Vung r149-r325: $($shift.Count) nguoi"

# ---- dem khop ----
$mCccd = 0; $mName = 0; $miss = New-Object System.Collections.ArrayList
foreach ($p in $shift) {
    $hit = $null
    if ($p.cccd -ne "" -and $traByCccd.ContainsKey($p.cccd)) { $hit = $traByCccd[$p.cccd]; $mCccd++ }
    elseif ($traByName.ContainsKey((NormName $p.ten))) { $hit = $traByName[(NormName $p.ten)]; $mName++ }
    if ($hit) { $p | Add-Member -NotePropertyName hoTra -NotePropertyValue $hit.hoTra -Force
                $p | Add-Member -NotePropertyName matchBy -NotePropertyValue $(if ($mCccd -gt 0 -and $p.cccd -ne "" -and $traByCccd.ContainsKey($p.cccd)) { "cccd" } else { "ten" }) -Force }
    else { [void]$miss.Add($p) }
}
L "Khop CCCD: $mCccd / theo ten: $mName / khong khop: $($miss.Count)"
L ""
L "KHONG KHOP ($($miss.Count)):"
foreach ($p in $miss) { L ("  dong {0}  ho{1}  {2}  cccd={3}" -f $p.row, $p.ho, $p.ten, $p.cccd) }

# ---- tao bang do hoTra -> STT ho To 1 tu cac dong TIN DANG (ngoai vung r149-325) ----
L ""
L "### Tao bang do hoTra -> STT ho To 1 tu dong khong bi chon ###"
$map = @{}
foreach ($p in ($t1 | Where-Object { $_.row -lt 149 -or $_.row -gt 325 })) {
    if (-not $p.hoTra) {
        $hit = $null
        if ($p.cccd -ne "" -and $traByCccd.ContainsKey($p.cccd)) { $hit = $traByCccd[$p.cccd] }
        elseif ($traByName.ContainsKey((NormName $p.ten))) { $hit = $traByName[(NormName $p.ten)] }
        if ($hit) { $p | Add-Member -NotePropertyName hoTra -NotePropertyValue $hit.hoTra -Force }
    }
    if ($p.hoTra -and $p.ho -gt 0) {
        $k = "$($p.hoTra)|$($p.ho)"
        if (-not $map.ContainsKey($k)) { $map[$k] = 0 }
        $map[$k]++
    }
}
# gom theo hoTra: co bao nhieu STT ho To 1 khac nhau?
$g = $map.Keys | ForEach-Object { $_.Split('|') } | Group-Object { $_[0] }
L "So hoTra xuat hien: $($g.Count)"
$conflict = @($g | Where-Object { $_.Count -gt 1 })
L "hoTra bi gan nhieu STT ho khac nhau: $($conflict.Count)"

# dem nguoi theo cap (hoTra, ho)
$pairs = @{}
foreach ($p in ($t1 | Where-Object { $_.hoTra -and ($_.row -lt 149 -or $_.row -gt 325) })) {
    $k = "$($p.hoTra)|$($p.ho)"; if (-not $pairs.ContainsKey($k)) { $pairs[$k] = 0 }; $pairs[$k]++
}
$pSorted = $pairs.Keys | Sort-Object { [int]($_.Split('|')[0]) }
L ""
L "### 25 cap (hoTra|hoTo1) dau, so nguoi ###"
foreach ($k in ($pSorted | Select-Object -First 25)) { L ("  {0,-16} {1} nguoi" -f $k, $pairs[$k]) }

# thu mapping: hoTra -> ho duy nhat
$hoOf = @{}; $amb = 0
foreach ($k in $pairs.Keys) {
    $a = $k.Split('|'); $h = [int]$a[1]
    if (-not $hoOf.ContainsKey($a[0])) { $hoOf[$a[0]] = $h } elseif ($hoOf[$a[0]] -ne $h) { $amb++ }
}
L ""
L "hoTra -> STT ho To 1: $($hoOf.Count) cap, $amb cap bi nhieu ho"
$recovered = 0; $noMap = New-Object System.Collections.ArrayList
foreach ($p in $shift) {
    if ($p.hoTra -and $hoOf.ContainsKey($p.hoTra)) { $recovered++ } else { [void]$noMap.Add($p) }
}
L "Vung r149-r325: khoi phuc duoc $recovered / $($shift.Count) ; khong co mapping: $($noMap.Count)"
L ""
L "### 30 nguoi vung chon dong sau khi khoi phuc ###"
foreach ($p in ($shift | Select-Object -First 30)) {
    $new = if ($p.hoTra -and $hoOf.ContainsKey($p.hoTra)) { $hoOf[$p.hoTra] } else { "?" }
    L ("  dong {0}  {1,-28} ho hien tai={2}  hoTra={3}  -> ho moi={4}" -f $p.row, $p.ten, $p.ho, $p.hoTra, $new)
}

Save-Utf8 "$OutDir\analyze-map.txt" $sb.ToString()
Write-Host "Xong. Xem $OutDir\analyze-map.txt"
