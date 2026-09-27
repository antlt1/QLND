@echo off
chcp 65001 >nul
REM ============================================================
REM  Chay web server QLND
REM  Bam kieu doi de chay. Bam Ctrl+C trong cua so nay de dung.
REM
REM  Cach dung:
REM    run-server.bat            -> cong 5173
REM    run-server.bat 5200       -> cong 5200
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

set "PORT=5173"
if not "%~1"=="" set "PORT=%~1"

REM neu chua co du lieu da gom thi chay build truoc
if not exist "output\to1-ho.json" (
  echo.
  echo   Chua co du lieu trong output\ - dang chay build...
  echo.
  call "%~dp0qlnd.cmd" build
  if errorlevel 1 (
    echo.
    echo   [LOI] Build that bai. Xem thong bao ben tren.
    echo.
    pause
    exit /b 1
  )
)

echo.
echo   Cong %PORT% dang chay? thu ket noi...
"%ND%\node.exe" -e "require('net').connect(%PORT%,'127.0.0.1').on('connect',function(){process.exit(0)}).on('error',function(){process.exit(1)})" 2>nul
if not errorlevel 1 (
  echo.
  echo   [LUU Y] Cong %PORT% da co chuong trinh chay.
  echo            Web dang o: http://localhost:%PORT%
  echo            Hay dung chuong trinh do, hoac chay lai voi cong khac:
  echo              run-server.bat 5200
  echo.
  pause
  exit /b 0
)

echo.
echo   ----------------------------------------------------
echo    TT:  http://localhost:%PORT%/
echo    Sua Excel dong 1-148:  http://localhost:%PORT%/fix.html
echo.
echo    Dung server:  Ctrl+C
echo   ----------------------------------------------------
echo.

"%ND%\node.exe" "%~dp0src\cli\serve.js" --port %PORT%

echo.
echo   Server da dung.
echo.
pause
endlocal
