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

$t1 = @(); $curA = ""; $curHo = 0
foreach ($r in (Read-Sheet $to1File 0)) {
    if ($r.Row -lt 3) { continue }
    $ten = Col $r 3; if ($ten -eq "") { continue }
    $a = Col $r 1
    if ($a -ne "") { $curA = $a; if ($a -match '^\d+$') { $curHo = [int]$a } }
    $c = Col $r 6
    $t1 += [pscustomobject]@{
        row = $r.Row; ten = $ten; ho = $curHo
        cccd = $(if ($c -match '^\d{12}$') { $c } else { "" })
    }
}
$tra = @(); $curHh = 0
foreach ($r in (Read-Sheet $traFile 0)) {
    if ($r.Row -lt 2) { continue }
    $ten = Col $r 2; if ($ten -eq "") { continue }
    if ((Col $r 7) -eq $CHU_HO) { $curHh++ }
    $cc = Col $r 5
    $tra += [pscustomobject]@{
        stt = (Col $r 1); ten = $ten; hoTra = $curHh
        cccd = $(if ($cc -match '^\d{12}$') { $cc } else { "" })
    }
}
$byC = @{}; foreach ($p in $tra) { if ($p.cccd -ne "" -and -not $byC.ContainsKey($p.cccd)) { $byC[$p.cccd] = $p } }
foreach ($p in $t1) { if ($p.cccd -ne "" -and $byC.ContainsKey($p.cccd)) { $p | Add-Member hoTra $byC[$p.cccd].hoTra -Force } }

$shift = @($t1 | Where-Object { $_.row -ge 149 -and $_.row -le 325 })
$trust = @($t1 | Where-Object { $_.row -lt 149 -or $_.row -gt 325 })

L "TOT1: $($t1.Count) nguoi | vung chon dong: $($shift.Count) | con lai: $($trust.Count)"
L "TRA: $($tra.Count) nguoi / 382 ho"
L ""
$shH = @($shift | Where-Object { $_.hoTra } | ForEach-Object { $_.hoTra } | Sort-Object -Unique)
$trH = @($trust | Where-Object { $_.hoTra } | ForEach-Object { $_.hoTra } | Sort-Object -Unique)
L "hoTra distinct: vung chon dong = $($shH.Count) | con lai = $($trH.Count)"
$overlap = @($shH | Where-Object { $trH -contains $_ })
L "hoTra dung chung (overlap) = $($overlap.Count)"
L ""
L "### Kiem tra 9 cap bi xung dot ###"
foreach ($h in $overlap) {
    $a = @($trust | Where-Object { $_.hoTra -eq $h })
    $b = @($shift | Where-Object { $_.hoTra -eq $h })
    L ""
    L "hoTra=$h  ->  TOT1 ho: $(($a | ForEach-Object { $_.ho } | Sort-Object -Unique) -join ',')"
    L "  [TOT1 con lai] $(($a | ForEach-Object { "r$($_.row)/ho$($_.ho) $($_.ten)" }) -join '  |  ')"
    L "  [vung chon dong] $(($b | ForEach-Object { "r$($_.row) $($_.ten)" }) -join '  |  ')"
    $full = @($tra | Where-Object { $_.hoTra -eq $h })
    L "  [FILE TRA ca ho] $(($full | ForEach-Object { "$($_.ten)$(if($_.ten -eq $b[0].ten){' <<<'})" }) -join '  |  ')"
}

L ""
L "### Phan bo so nguoi moi hoTra trong FILE TRA ###"
$g = $tra | Group-Object hoTra
$g | Group-Object Count | Sort-Object { [int]$_.Name } | ForEach-Object { L ("  hoTra co {0,2} nguoi: {1} ho" -f $_.Name, $_.Count) }

L ""
L "### Cac nguoi TOT1 (con lai) co hoTra bi trung nhau ###"
$tg = $trust | Where-Object { $_.hoTra } | Group-Object hoTra | Where-Object { $_.Count -gt 1 }
L "so hoTra bi >1 nguoi TOT1: $($tg.Count)"
foreach ($x in ($tg | Sort-Object Count -Descending | Select-Object -First 15)) {
    L "  hoTra=$($x.Name) n=$($x.Count) ho:$(($x.Group | ForEach-Object { $_.ho } | Sort-Object -Unique) -join ',')  $(($x.Group | ForEach-Object { $_.ten }) -join ' | ')"
}

Save-Utf8 "$OutDir\analyze-map2.txt" $sb.ToString()
Write-Host "Xong."
