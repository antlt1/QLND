param(
    [string]$SrcDir = "C:\Users\Administrator\Downloads\An\QLND\source\fileexcel",
    [string]$OutDir = "C:\Users\Administrator\Downloads\An\QLND\output"
)
$ErrorActionPreference = "Stop"
. "$PSScriptRoot\xlsx-lib.ps1"

# =====================================================================
# DOC TOT 1 -> danh sach nguoi (chua gan ho)
# =====================================================================
$to1File = $null; $traFile = $null; $traAFile = $null
foreach ($f in (Get-ChildItem -LiteralPath $SrcDir -Filter "*.xlsx")) {
    if ($f.Name -like "* 1.xlsx" -and $f.Length -gt 100000) { $to1File = $f.FullName }
    if ($f.Name -like "*tra.xlsx") {
        if ($f.Length -lt 140000) { $traFile = $f.FullName } else { $traAFile = $f.FullName }
    }
}
if (-not $to1File) { throw "Khong tim thay file To 1 trong $SrcDir" }

$sbLog = New-Object System.Text.StringBuilder
function L($s) { [void]$sbLog.AppendLine($s) }

$issues = New-Object System.Collections.ArrayList
function Flag($row, $ho, $ten, $code, $msg) {
    [void]$issues.Add([pscustomobject]@{ row = $row; ho = $ho; ten = $ten; code = $code; moTa = $msg })
}

# ---------- Pass 1: doc tung dong ----------
$people = New-Object System.Collections.ArrayList
$curHoFwd = 0; $curARaw = ""
foreach ($r in (Read-Sheet $to1File 0)) {
    if ($r.Row -lt 3) { continue }
    $ten = Col $r 3
    if ($ten -eq "") { continue }
    $aRaw = Col $r 1
    if ($aRaw -ne "") { $curARaw = $aRaw; if ($aRaw -match '^\d+$') { $curHoFwd = [int]$aRaw } }
    $rowShifted = $false

    # --- moi quan he: dung vi tri dung theo tung block ---
    $quanHe = $null; $quanHeCol = $null
    if ($REL -contains (Col $r 9))     { $quanHe = Col $r 9; $quanHeCol = "I" }
    elseif ($REL -contains (Col $r 8)) { $quanHe = Col $r 8; $quanHeCol = "H" }
    elseif ($REL -contains (Col $r 7)) { $quanHe = Col $r 7; $quanHeCol = "G" }
    if (-not $quanHe) { Flag $r.Row $aRaw $ten "THIEU_QUAN_HE" "Khong co 'chu ho' / 'thanh vien'" }

    # --- ngay sinh / ngay cap ---
    $eRaw = Col $r 5
    if ($eRaw -match '^\d{12}$' -and (Col $r 6) -eq "") {
        $rowShifted = $true
        Flag $r.Row $aRaw $ten "DONG_BI_LECH_COT" "CCCD '$eRaw' nam o cot E, 'moi quan he' o cot G (thay vi F va I) - da tam khoi phuc"
    }
    $ns = $null
    if ($eRaw -ne "" -and -not $rowShifted) {
        $ns = Ser2Date $eRaw; if (-not $ns) { $ns = Txt2Date $eRaw }
        if (-not $ns) { Flag $r.Row $aRaw $ten "NGAY_SINH_KHONG_DOC_DUOC" "Gia tri '$eRaw' khong phai ngay sinh hop le" }
    } elseif (-not $rowShifted) { Flag $r.Row $aRaw $ten "THIEU_NGAY_SINH" "Cot E trong" }
    $nc = $null
    $gRaw = Col $r 7
    if ($gRaw -ne "" -and -not $rowShifted) {
        $nc = Ser2Date $gRaw; if (-not $nc) { $nc = Txt2Date $gRaw }
        if (-not $nc) { Flag $r.Row $aRaw $ten "NGAY_CAP_KHONG_DOC_DUOC" "Gia tri '$gRaw'" }
    }

    # --- tuoi (chi block 1) ---
    $tuoi = $null
    if ($r.Row -le 328) {
        $hRaw = Col $r 8
        if ($hRaw -match '^\d+$') { $tuoi = [int]$hRaw }
        elseif ($hRaw -ne "") { Flag $r.Row $aRaw $ten "TUOI_KHONG_PHAI_SO" "Gia tri '$hRaw'" }
    }

    # --- cccd ---
    $fRaw = Col $r 6; $cccd = ""
    if ($rowShifted) { $cccd = $eRaw }
    elseif ($fRaw -match '^\d{12}$') { $cccd = $fRaw }
    elseif ($fRaw -ne "") { Flag $r.Row $aRaw $ten "CCCD_KHONG_DUNG_DOI" "Gia tri '$fRaw' khong phai 12 so" }
    if ($cccd -eq "") { Flag $r.Row $aRaw $ten "THIEU_CCCD" "Khong co so dinh danh" }

    # --- ma bao hiem ---
    $maBH = Col $r 4
    $khongBH = $false
    if ($maBH -eq $KHONG_CO) { $maBH = ""; $khongBH = $true }
    if ($maBH -eq "") { Flag $r.Row $aRaw $ten "THIEU_MA_BH" "Khong co ma bao hiem hoac 'khong co'" }

    [void]$people.Add([pscustomobject][ordered]@{
        row = $r.Row
        block = $(if ($r.Row -le 328) { 1 } else { 2 })
        rowShifted = $rowShifted
        cotARaw = $aRaw            # gia tri o cot A tai dong nay
        hoGoc = $curHoFwd          # so ho forward-fill (gia tri goc)
        coCotA = ($aRaw -ne "")    # dong nay co so ho rieng -> tin cay        hoTen = $ten
        sttNguoiGoc = (Col $r 2)
        quanHe = $quanHe; quanHeCol = $quanHeCol
        maBH = $maBH; khongBH = $khongBH
        ngaySinh = $ns; ngaySinhRaw = $eRaw
        cccd = $cccd; ngayCap = $nc; tuoi = $tuoi
        dienThoai = (Col $r 10); duHoc = (Col $r 11); layChong = (Col $r 12)
        xkld = (Col $r 13); ngheNghiep = (Col $r 14); ghiChu = (Col $r 25)
        datDai = [pscustomobject]@{
            lua      = [pscustomobject]@{ dienTich = (Col $r 15) }
            vuon     = [pscustomobject]@{ cay = (Col $r 16); soLuong = (Col $r 17); dienTich = (Col $r 18) }
            rauMau   = [pscustomobject]@{ tenRau = (Col $r 19); dienTich = (Col $r 20) }
            ao       = [pscustomobject]@{ tenNuoi = (Col $r 21); soLuong = (Col $r 22) }
            chanNuoi = [pscustomobject]@{ giaXuc = (Col $r 23); giaCam = (Col $r 24) }
        }
        canhBao = $(if ($rowShifted) { @("DONG_BI_LECH_COT_DA_KHOI_PHUCCCD") } else { @() })
        hoTra = $null; matchBy = ""; sttHo = 0; sttNguoi = 0; vaiTro = ""
        nguonHo = "cot-A"
    })
}
L "To 1: $($people.Count) nguoi doc duoc"
$needRec = @($people | Where-Object { -not $_.coCotA })
L "Dong thanh vien (khong co STT ho rieng o cot A): $($needRec.Count)"

# ---------- Pass 2: doi chieu CCCD voi file tra ----------
$traOut = @()
if ($traFile) {
    $tra = Read-Sheet $traFile 0
    $curHh = 0
    foreach ($r in $tra) {
        if ($r.Row -lt 2) { continue }
        $ten = Col $r 2; if ($ten -eq "") { continue }
        if ((Col $r 7) -eq $CHU_HO) { $curHh++ }
        $cc = Col $r 5
        $traOut += [pscustomobject][ordered]@{
            stt = (Col $r 1); hoTen = $ten
            soDinhDanh = $(if ((Col $r 3) -eq $KHONG_CO) { "" } else { (Col $r 3) })
            ngaySinh = (Txt2Date (Col $r 4))
            cccd = $(if ($cc -match '^\d{12}$') { $cc } else { "" })
            ngayCap = (Col $r 6); quanHe = (Col $r 7)
            hoTraStt = $curHh
        }
    }
    L "File tra: $($traOut.Count) nguoi / $curHh ho"
}
$byCccd = @{}; $byName = @{}
foreach ($p in $traOut) {
    if ($p.cccd -ne "" -and -not $byCccd.ContainsKey($p.cccd)) { $byCccd[$p.cccd] = $p }
    $nk = NormName $p.hoTen
    if ($nk -ne "" -and -not $byName.ContainsKey($nk)) { $byName[$nk] = $p }
}
foreach ($p in $people) {
    $hit = $null
    if ($p.cccd -ne "" -and $byCccd.ContainsKey($p.cccd)) { $hit = $byCccd[$p.cccd]; $p.matchBy = "CCCD" }
    elseif ($byName.ContainsKey((NormName $p.hoTen))) { $hit = $byName[(NormName $p.hoTen)]; $p.matchBy = "ten" }
    if ($hit) { $p.hoTra = $hit.hoTraStt }
}
$mC = @($needRec | Where-Object { $_.matchBy -eq "CCCD" }).Count
$mT = @($needRec | Where-Object { $_.matchBy -eq "ten" }).Count
$mN = $needRec.Count - $mC - $mT
L "Khop CCCD: $mC / theo ten: $mT / khong khop: $mN"

# ---------- Pass 3: dem CCCD trung (dung de goi xung dot) ----------
$dupRows = @{}
$grp = $people | Where-Object { $_.cccd -ne "" } | Group-Object cccd
foreach ($g in $grp) {
    if ($g.Count -lt 2) { continue }
    $rows = ($g.Group | ForEach-Object { $_.row }) -join ', '
    foreach ($m in $g.Group) {
        $dupRows[$m.row] = $true
        $m.canhBao = @($m.canhBao) + "CCCD_TRUNG_VOI_DONG_KHAC"
        Flag $m.row $m.hoGoc $m.hoTen "CCCD_TRUNG" "CCCD $($g.Name) xuat hien o dong $rows"
    }
}

# ---------- Pass 4: bang do hoTra -> STT ho To 1 ----------
# Chi dung cac dong TIN DAY (co so ho rieng o cot A) va loai dong bi trung CCCD khi co xung dot
$cand = @{}   # hoTra -> hashtable { hoSo = soNguoi }
foreach ($p in ($people | Where-Object { $_.coCotA -and $_.hoTra -and $_.hoGoc -gt 0 })) {
    if (-not $cand.ContainsKey($p.hoTra)) { $cand[$p.hoTra] = @{} }
    $k = [string]$p.hoGoc
    if (-not $cand[$p.hoTra].ContainsKey($k)) { $cand[$p.hoTra][$k] = 0 }
    $cand[$p.hoTra][$k]++
}
$hoMap = @{}; $ambig = @{}
foreach ($h in $cand.Keys) {
    $keys = @($cand[$h].Keys)
    if ($keys.Count -eq 1) { $hoMap[$h] = [int]$keys[0]; continue }
    # bo dong bi trung CCCD, xem lai
    $clean = @()
    foreach ($k in $keys) {
        $rowsOk = @($people | Where-Object { $_.coCotA -and $_.hoTra -eq $h -and $_.hoGoc -eq [int]$k -and -not $dupRows.ContainsKey($_.row) })
        if ($rowsOk.Count -gt 0) { $clean += $k }
    }
    if ($clean.Count -eq 1) {
        $hoMap[$h] = [int]$clean[0]
        L "  hoTra=${h}: bo qua dong trung CCCD, chon ho $($clean[0])"
    } else {
        $ambig[$h] = ($keys -join '/')
        L "  hoTra=${h}: XUNG DOT $((($keys | ForEach-Object { "ho$_" }) -join ' / ')) - bo qua"
    }
}
L "Bang do hoTra -> STT ho: $($hoMap.Count) cap hop le, $($ambig.Count) cap xung dot"

# ---------- Pass 5: gan lai STT ho cho dong bi chon ----------
$fixed = 0; $unfixed = 0
foreach ($p in $needRec) {
    if ($p.hoTra -and $hoMap.ContainsKey($p.hoTra)) {
        $p.sttHo = $hoMap[$p.hoTra]
        $p.nguonHo = "khoi-phuc-CCCD"
        $p.canhBao = @($p.canhBao) + "DA_KHOI_PHUCC_HO_TU_DS_TRUONG_BINH"
        $fixed++
    } else {
        $p.sttHo = $p.hoGoc          # giu nguyen forward-fill
        $p.nguonHo = "cot-A(forward)"
        $p.canhBao = @($p.canhBao) + "CHUA_KHOI_PHU_COT_THAU_VI_TRA_DS_TRUONG_BINH"
        if ($p.matchBy -eq "") { Flag $p.row $p.hoGoc $p.hoTen "KHONG_KHOP_DS_TRA" "Khong tim thay trong danh sach Trau Binh theo CCCD hay ten" }
        elseif ($ambig.ContainsKey($p.hoTra)) { Flag $p.row $p.hoGoc $p.hoTen "HO_TRA_XUNG_DOT" "Ho trong DS tra #$($p.hoTra) co bien theo nhieu STT ho To 1 ($($ambig[$p.hoTra]))" }
        else { Flag $p.row $p.hoGoc $p.hoTen "HO_TRA_CHUA_CO_MAPPING" "Ho trong DS tra #$($p.hoTra) chua co nguoi nao khop de xac dinh STT ho To 1" }
        $unfixed++
    }
}
L "Khoi phuc STT ho: $fixed dong | giu nguyen: $unfixed dong"
foreach ($p in ($people | Where-Object { $_.coCotA })) { $p.sttHo = $p.hoGoc; $p.nguonHo = "cot-A" }

# ---------- Pass 6: gom theo ho ----------
# Key: dong chu ho (cot A co gia tri) -> "N<so>" neu la so, "A<row>" neu cot A rac (gia tri goc, tach rieng)
#       dong thanh vien -> "N<sttHo>" (forward-fill hoac da khoi phuc)
$households = New-Object System.Collections.ArrayList
$hhIdx = @{}
foreach ($p in $people) {
    if ($p.coCotA) { $key = $(if ($p.cotARaw -match '^\d+$') { "N$($p.cotARaw)" } else { "A$($p.row)" }) }
    else { $key = "N$($p.sttHo)" }
    if (-not $hhIdx.ContainsKey($key)) {
        $hh = [pscustomobject]@{
            key = $key
            sttHo = $(if ($p.coCotA -and $p.cotARaw -match '^\d+$') { [int]$p.cotARaw } else { 0 })
            sttHoGoc = $(if ($p.coCotA) { $p.cotARaw } else { "" })
            cotARaw = $(if ($p.coCotA) { $p.cotARaw } else { "" })
            nguoi = (New-Object System.Collections.ArrayList)
            chuHo = $null; chuHoRow = 0; canhBao = @(); goiY = ""
            rowMin = $p.row; rowMax = $p.row
            khung = "KHONG_XAC_DINH"; soNguoi = 0; id = ""
        }
        [void]$households.Add($hh)
        $hhIdx[$key] = $households.Count - 1
    }
    $hh = $households[$hhIdx[$key]]
    [void]$hh.nguoi.Add($p)
    if ($p.row -lt $hh.rowMin) { $hh.rowMin = $p.row }
    if ($p.row -gt $hh.rowMax) { $hh.rowMax = $p.row }
    if ($p.quanHe -eq $CHU_HO -and -not $hh.chuHo) { $hh.chuHo = $p.hoTen; $hh.chuHoRow = $p.row }
    if ($p.quanHe -eq $CHU_HO) {
        $d = @($hh.nguoi | Where-Object { $_.row -lt $p.row -and $_.quanHe -eq $CHU_HO })
        if ($d.Count -gt 0) {
            Flag $p.row $hh.sttHo $p.hoTen "TRONG_HO_HAI_CHU_HO" "Nhom co $($d.Count + 1) dong danh 'chu ho' (dong $($hh.chuHoRow) va dong $($p.row))"
        }
    }
}
# sap xep theo dong chu ho
$sorted = @($households | Sort-Object { if ($_.chuHoRow -gt 0) { $_.chuHoRow } else { $_.rowMin } })
$households = New-Object System.Collections.ArrayList
foreach ($h in $sorted) { [void]$households.Add($h) }

# ---------- Pass 7: STT nguoi lien tuc, vai tro, khung, canh bao ----------
$stt = 0; $hhNo = 0
foreach ($h in $households) {
    $hhNo++
    $h | Add-Member -NotePropertyName no -NotePropertyValue $hhNo -Force
    $first = $true
    foreach ($m in ($h.nguoi | Sort-Object row)) {
        $stt++
        $m.sttNguoi = $stt
        if ($first) {
            $m.vaiTro = $(if ($m.quanHe -eq $CHU_HO) { "chuHo" } else { "thanhVien" })
            $first = $false
        } else {
            $m.vaiTro = "thanhVien"
            if ($m.quanHe -eq $CHU_HO) { $m.canhBao = @($m.canhBao) + "DONG_DAU_KHONG_PHAI_CHU_HO" }
        }
    }
    $n = 0
    $khung = "KHONG_XAC_DINH"
    if ($h.sttHoGoc -match '^\d+$') {
        $n = [int]$h.sttHoGoc
        if     ($n -le 27)  { $khung = "1_DA_FIX" }
        elseif ($n -le 78)  { $khung = "2_CHUA_FIX_THANH_VIEN_BI_CHON" }
        elseif ($n -le 141) { $khung = "3_CHUA_CAP_NHAT" }
        else                { $khung = "NGOAI_PHU" }
    }
    $h.khung = $khung; $h.sttHo = $n
    $h.soNguoi = $h.nguoi.Count
    $h.id = "HO-" + $h.rowMin
    $nFix = @($h.nguoi | Where-Object { $_.nguonHo -eq "khoi-phuc-CCCD" }).Count
    $w = New-Object System.Collections.ArrayList
    if ($h.cotARaw -and $h.cotARaw -notmatch '^\d+$') { [void]$w.Add("STT_HO_RAC") }
    if ($h.soNguoi -eq 1) { [void]$w.Add("HO_CHI_CO_1_NGUOI") }
    if ($khung -eq "2_CHUA_FIX_THANH_VIEN_BI_CHON") { [void]$w.Add("CAN_THEM_THANH_VIEN_TU_DS_TRUONG_BINH") }
    if ($khung -eq "3_CHUA_CAP_NHAT") { [void]$w.Add("CHUA_CAP_NHAT_TU_DS_TRUONG_BINH") }
    if (-not $h.chuHo) { [void]$w.Add("THIEU_CHU_HO") }
    if ($nFix -gt 0) { [void]$w.Add("DA_KHOI_PHU_$nFix_THANH_VIEN") }
    foreach ($m in $h.nguoi) { foreach ($c in $m.canhBao) { [void]$w.Add("NGUOI_R$c") } }
    $h.canhBao = @($w)

    # goi y cho nhom sinh tu cot A rac
    if ($h.cotARaw -and $h.cotARaw -notmatch '^\d+$') {
        $prev = $null
        for ($k = $hhNo - 2; $k -ge 0; $k--) { if ($households[$k].cotARaw -match '^\d+$') { $prev = $households[$k]; break } }
        if ($prev) { $h.goiY = "Cot A dong $($h.rowMin)-$($h.rowMax) = '$($h.cotARaw)'; theo vi tri thi thuoc ho $($prev.sttHoGoc) (bat dau dong $($prev.rowMin)). Can kiem tra lai." }
    }
}

# ---------- GHI FILE ----------
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

$hhJson = @($households | ForEach-Object {
    [ordered]@{
        id = $_.id; no = $_.no; sttHo = $_.sttHo; cotARaw = $_.cotARaw
        khung = $_.khung; chuHo = $_.chuHo
        rowMin = $_.rowMin; rowMax = $_.rowMax; soNguoi = $_.soNguoi
        canhBao = $_.canhBao; goiY = $_.goiY
        thanhVien = @($_.nguoi | Sort-Object sttNguoi | ForEach-Object {
            [ordered]@{
                sttNguoi = $_.sttNguoi; row = $_.row; block = $_.block
                hoTen = $_.hoTen; quanHe = $_.quanHe; vaiTro = $_.vaiTro
                maBH = $_.maBH; khongBH = $_.khongBH
                ngaySinh = $_.ngaySinh; cccd = $_.cccd; ngayCap = $_.ngayCap
                tuoi = $_.tuoi; dienThoai = $_.dienThoai
                duHoc = $_.duHoc; layChong = $_.layChong; xkld = $_.xkld
                ngheNghiep = $_.ngheNghiep; ghiChu = $_.ghiChu
                cotARaw = $_.cotARaw; hoGoc = $_.hoGoc
                hoTra = $_.hoTra; matchBy = $_.matchBy
                nguonHo = $_.nguonHo; canhBao = $_.canhBao
                datDai = $_.datDai
            }
        })
    }
})
$doc = [ordered]@{
    nguon = (Split-Path $to1File -Leaf)
    ap = (U "1EA4 70 72 1B0D 1EE3 6E 67 20 42 1ECB 6E 68")
    to = 1
    xuatAt = (Get-Date).ToString("s")
    tongNguoi = $people.Count; tongHo = $households.Count
    ghiChu = "Cot A = STT ho (chi ghi o dong chu ho). STT nguoi danh lai lien tuc 1..N. Khong xoa dong nao."
    quyTac = [ordered]@{
        cotA      = "STT ho, chi ghi o dong chu ho"
        khoiPhuc  = "Dong khong co STT ho rieng duoc gan lai bang cach doi chieu CCCD/ten voi Danh sach Trau Binh, dua tren bang do ho cua nhung dong con lai."
    }
    ho = $hhJson
}
Save-Utf8 "$OutDir\to1-ho.json" ($doc | ConvertTo-Json -Depth 12)

$flat = New-Object System.Collections.ArrayList
foreach ($h in $households) {
    foreach ($m in ($h.nguoi | Sort-Object sttNguoi)) {
        [void]$flat.Add([pscustomobject][ordered]@{
            sttHo = $h.sttHo; cotARaw = $h.cotARaw; khung = $h.khung
            sttNguoi = $m.sttNguoi; rowExcel = $m.row; block = $m.block
            hoTen = $m.hoTen; quanHe = $m.quanHe; vaiTro = $m.vaiTro
            maBH = $m.maBH; ngaySinh = $m.ngaySinh; cccd = $m.cccd
            ngayCap = $m.ngayCap; tuoi = $m.tuoi; dienThoai = $m.dienThoai
            ngheNghiep = $m.ngheNghiep; duHoc = $m.duHoc; layChong = $m.layChong
            xkld = $m.xkld; ghiChu = $m.ghiChu
            hoTra = $m.hoTra; nguonHo = $m.nguonHo
            luaDT = $m.datDai.lua.dienTich; vuonCay = $m.datDai.vuon.cay
            vuonSL = $m.datDai.vuon.soLuong; vuonDT = $m.datDai.vuon.dienTich
            rauTen = $m.datDai.rauMau.tenRau; rauDT = $m.datDai.rauMau.dienTich
            aoTen = $m.datDai.ao.tenNuoi; aoSL = $m.datDai.ao.soLuong
            giaXuc = $m.datDai.chanNuoi.giaXuc; giaCam = $m.datDai.chanNuoi.giaCam
            canhBao = (($m.canhBao | Sort-Object -Unique) -join " | ")
        })
    }
}
$flat | Export-Csv -LiteralPath "$OutDir\to1-nguoi.csv" -NoTypeInformation -Encoding UTF8
$issues | Export-Csv -LiteralPath "$OutDir\to1-loi.csv" -NoTypeInformation -Encoding UTF8
Save-Utf8 "$OutDir\to1-loi.json" ($issues | ConvertTo-Json -Depth 5)
if ($traOut.Count -gt 0) {
    $traOut | Export-Csv -LiteralPath "$OutDir\tra-cuoc.csv" -NoTypeInformation -Encoding UTF8
    Save-Utf8 "$OutDir\tra-cuoc.json" ($traOut | ConvertTo-Json -Depth 6)
}

$sum = @()
foreach ($k in @("1_DA_FIX","2_CHUA_FIX_THANH_VIEN_BI_CHON","3_CHUA_CAP_NHAT","NGOAI_PHU","KHONG_XAC_DINH")) {
    $sel = @($households | Where-Object { $_.khung -eq $k })
    if (-not $sel.Count) { continue }
    $n = 0; foreach ($s in $sel) { $n += $s.soNguoi }
    $sum += [pscustomobject]@{ khung = $k; soHo = $sel.Count; soNguoi = $n }
}
$sum | Format-Table -AutoSize | Out-String -Width 120 | ForEach-Object { L $_ }
L ""
L "TONG: $($households.Count) ho / $($people.Count) nguoi / $($issues.Count) canh bao"
$issues | Group-Object code | Sort-Object Count -Descending | ForEach-Object { L ("  {0,-30} {1}" -f $_.Name, $_.Count) }
Save-Utf8 "$OutDir\extract.log" $sbLog.ToString()
Write-Host "XONG -> $OutDir"
