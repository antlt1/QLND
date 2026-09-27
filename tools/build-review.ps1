param(
    [string]$OutDir = "C:\Users\Administrator\Downloads\An\QLND\output"
)
$ErrorActionPreference = "Stop"
$tpl = Get-Content -LiteralPath "$PSScriptRoot\review.template.html" -Raw -Encoding UTF8

function J($p) {
    $raw = Get-Content -LiteralPath $p -Raw -Encoding UTF8
    # thoat dau "</script>" trong du lieu de khong lam hong the HTML
    return $raw -replace '</', '<\/'
}

$to1  = J "$OutDir\to1-ho.json"
$loi  = J "$OutDir\to1-loi.json"
$tra  = J "$OutDir\tra-cuoc.json"

$html = $tpl.Replace('__TO1__', $to1).Replace('__LOI__', $loi).Replace('__TRA__', $tra)
$enc = New-Object System.Text.UTF8Encoding($true)
$dest = "$OutDir\review.html"
[System.IO.File]::WriteAllText($dest, $html, $enc)
Write-Host ("Da tao {0} ({1:N0} bytes)" -f $dest, (Get-Item $dest).Length)
