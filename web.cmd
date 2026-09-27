@echo off
chcp 65001 >nul
REM ============================================================
REM  Mo web xem danh sach ho - ap Truong Binh, To 1
REM  Bam kieu doi de mo. Bam Ctrl+C trong cua so nay de dung.
REM ============================================================
setlocal
cd /d "%~dp0"

set "ND=%USERPROFILE%\.qlnd\node"
if not exist "%ND%\node.exe" (
  echo.
  echo   [LOI] Khong tim thay Node portable tai "%ND%"
  echo         Xem muc 2 trong README.md de cai lai.
  echo.
  pause
  exit /b 1
)
set "PATH=%ND%;%PATH%"

REM cong mac dinh 5173, co the doi:  web.cmd 5200
if not "%~1"=="" set "PORT=%~1"

echo.
echo   Dang doc du lieu, vui long cho...
call "%~dp0qlnd.cmd" build
if errorlevel 1 (
  echo.
  echo   [LOI] Doc du lieu that bai. Xem thong bao ben tren.
  echo.
  pause
  exit /b 1
)

echo.
echo   Dang khoi dong web tren cong %PORT%...

REM bat server o che do nen
start "QLND web" /min "%ND%\node.exe" "%~dp0src\cli\serve.js" --port %PORT%

REM cho server co time len
set /a i=0
:wait
timeout /t 1 /nobreak >nul
set /a i+=1
"%ND%\node.exe" -e "require('net').connect(%PORT%,'127.0.0.1').on('connect',function(){process.exit(0)}).on('error',function(){process.exit(1)})" 2>nul
if not errorlevel 1 goto open
if %i% lss 20 goto wait

echo.
echo   [LOI] Web khong khoi dong duoc sau 20 giay.
echo         Co the cong %PORT% dang duoc chuong trinh khai chiem.
echo         Thu lai voi cong khac:  web.cmd 5200
echo.
pause
exit /b 1

:open
start "" "http://localhost:%PORT%"
echo   Da mo trinh duyet:  http://localhost:%PORT%
echo.
echo   ----------------------------------------------------
echo    TT: http://localhost:%PORT%
echo    - Bam vao mot ho de xem chi tiet nguoi trong ho
echo    - Tim theo ten / STT ho / CCCD / so dien thoai
echo    - Loc theo khoi du lieu
echo.
echo    Dung web:  taskkill /FI "WINDOWTITLE eq QLND web*" /T
echo   ----------------------------------------------------
echo.
pause
endlocal
