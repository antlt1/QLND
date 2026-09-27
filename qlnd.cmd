@echo off
REM launcher cho QLND - tu them Node portable vao PATH roi chay lenh
REM usage: qlnd.cmd <lenh>   vi du: qlnd.cmd extract | excel | xep | fixroute | serve | test
setlocal

REM Neu bam dup (khong phai go tu terminal) thi giu cua so lai de xem ket qua.
REM Dat QLND_NO_PAUSE=1 de tat (dung khi chay tu script/CI).
set "QLND_PAUSE=0"
if not "%QLND_NO_PAUSE%"=="1" echo %cmdcmdline% | find /i "%~nx0" >nul 2>&1 && set "QLND_PAUSE=1"

set "ND=%USERPROFILE%\.qlnd\node"
if not exist "%ND%\node.exe" (
  echo [LOI] Khong tim thay Node portable tai "%ND%".
  echo       Chay lai: node src\setup-node.js
  goto ketqua
)
set "PATH=%ND%;%ND%\node_modules\npm\bin;%PATH%"

if "%~1"=="" goto help
if /i "%~1"=="help" goto help
if /i "%~1"=="-h" goto help
if /i "%~1"=="--help" goto help

REM "test" nam ngoai src/cli
if /i "%~1"=="test" goto chaytest

set "CMD=%~1"
shift
%ND%\node.exe "%~dp0src\cli\%CMD%.js" %1 %2 %3 %4 %5 %6 %7 %8 %9
goto ketqua

:chaytest
%ND%\node.exe "%~dp0test\run.js"
goto ketqua

:help
echo.
echo   QLND - quan ly danh sach ho can bo ap Truong Binh
echo.
echo   Cach dung:
echo     qlnd.cmd extract     Doc file Excel, gom ho, khoi phuc STT ho  -^> output\*.json
echo     qlnd.cmd excel       Xuat file Excel sach                      -^> output\*.xlsx
echo     qlnd.cmd build       Chay ca hai
echo     qlnd.cmd xep         Sap xep lai tung ho: chu ho roi den thanh vien
echo     qlnd.cmd fixroute    Xem/sua tung dong 1..148 cua file To 1
echo     qlnd.cmd serve       Mo web xem du lieu tai http://localhost:5173
echo     qlnd.cmd test        Chay test
echo.
echo   Fix route (duyet tung dong):
echo     qlnd.cmd fixroute                 Chi kiem tra dong 1..148, khong sua
echo     qlnd.cmd fixroute --sua           Sua nhung gi chan chan, xuat file moi
echo     qlnd.cmd fixroute --from 1 --to 148 --sua
echo.
echo   Xep lai danh sach (chu ho + thanh vien):
echo     qlnd.cmd xep
echo     qlnd.cmd xep --ap "Truong Binh" --out D:\ketqua
echo.
echo   Tu chon thu muc khac:
echo     qlnd.cmd extract --src D:\data --out D:\ketqua
echo.

:ketqua
REM %ERRORLEVEL% phai lay o day, sau khi lenh vua chay xong
if "%QLND_PAUSE%"=="1" (
  echo.
  pause
)
endlocal & exit /b %ERRORLEVEL%
