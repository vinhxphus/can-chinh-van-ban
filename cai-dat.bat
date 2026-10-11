@echo off
rem Cai dat / cap nhat add-in "Can chinh van ban" cho Word tren may nay.
rem Chi can chay file nay (khong can quyen Administrator). Tai ve tai:
rem   https://vinhxphus.github.io/can-chinh-van-ban/cai-dat.bat
setlocal

set "URL=https://vinhxphus.github.io/can-chinh-van-ban/manifest-prod.xml"
set "DIR=C:\WordAddin"
set "MANIFEST=%DIR%\manifest-prod.xml"
set "ADDIN_ID={b6f0a1a2-4c1e-4a7a-9f0a-3d9e0c1e2f10}"
set "REGKEY=HKCU\Software\Microsoft\Office\16.0\WEF\Developer"

echo ==============================================
echo   Cai dat add-in "Can chinh van ban" cho Word
echo ==============================================
echo.

rem 1. Word phai dong de xoa duoc bo nho dem (cache) add-in cu.
:check_word
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
  echo [!] Word dang mo. Hay LUU tai lieu va DONG HET Word, roi nhan phim bat ky de tiep tuc...
  pause >nul
  goto check_word
)

rem 2. Tai manifest moi nhat tu GitHub Pages; khong tai duoc thi dung file manifest-prod.xml dat canh file nay.
if not exist "%DIR%" mkdir "%DIR%"
echo [1/3] Tai manifest moi nhat...
rem Luu nguyen file tai ve (giu dung ma UTF-8 tieng Viet), kiem tra dung add-in roi moi thay file cu.
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { Invoke-WebRequest -UseBasicParsing -Uri '%URL%' -TimeoutSec 30 -OutFile '%MANIFEST%.tmp'; if (-not (Select-String -Path '%MANIFEST%.tmp' -Pattern 'b6f0a1a2-4c1e-4a7a-9f0a-3d9e0c1e2f10' -Quiet)) { exit 2 }; Move-Item -Force '%MANIFEST%.tmp' '%MANIFEST%'; exit 0 } catch { exit 1 }"
if errorlevel 1 (
  if exist "%MANIFEST%.tmp" del "%MANIFEST%.tmp"
  if exist "%~dp0manifest-prod.xml" (
    copy /Y "%~dp0manifest-prod.xml" "%MANIFEST%" >nul
    echo       Khong tai duoc tu mang - dung file manifest-prod.xml dat canh file cai dat.
  ) else (
    echo [X] Khong tai duoc manifest va khong co file manifest-prod.xml canh file cai dat.
    echo     Kiem tra ket noi mang roi chay lai.
    pause
    exit /b 1
  )
) else (
  echo       Da luu %MANIFEST%
)

rem 3. Dang ky add-in voi Word (xoa ban dang ky ban dev/localhost cung ID neu co).
echo [2/3] Dang ky add-in voi Word...
reg delete "%REGKEY%" /v "b6f0a1a2-4c1e-4a7a-9f0a-3d9e0c1e2f10" /f >nul 2>&1
reg add "%REGKEY%" /v "%ADDIN_ID%" /t REG_SZ /d "%MANIFEST%" /f >nul
if errorlevel 1 (
  echo [X] Khong ghi duoc registry.
  pause
  exit /b 1
)

rem 4. Xoa cache add-in de Word nap ban moi.
echo [3/3] Xoa bo nho dem add-in cu...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" rmdir /S /Q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef"

echo.
echo Xong! Mo Word, tren tab Home (Trang chu) bam nut "Can chinh VB".
echo Lan sau chi can chay lai file nay de cap nhat khi co ban moi.
echo.
pause
