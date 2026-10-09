@echo off
rem ===========================================================================
rem  Excel Merger - build (for developers).  KEEP THIS FILE ASCII-ONLY + CRLF.
rem    build.bat            npm install, UI build, Excel runner, ExcelMerger.exe
rem    build.bat nopause    same, without the final "pause" (scripts / CI)
rem  Output: dist\index.html, server\excel\excel-runner.ps1, ExcelMerger.exe
rem  Hand-out folder + zip: npm run release  (= node scripts\make-release.mjs; release\ExcelMerger and release\ExcelMerger_yyyyMMdd.zip, exe is NOT rebuilt)
rem  Plan B for PCs that block unsigned exes: npm run build:exe -- --layout runtime
rem ===========================================================================
setlocal
pushd "%~dp0" || goto :nofolder
title Excel Merger - build

where node >nul 2>nul
if errorlevel 1 goto :nonode
set "NODE_MAJOR="
for /f "tokens=1 delims=v." %%v in ('node --version 2^>nul') do if not defined NODE_MAJOR set "NODE_MAJOR=%%v"
if not defined NODE_MAJOR goto :nonode
if %NODE_MAJOR% LSS 22 goto :oldnode
echo [INFO] Node.js version (this node.exe is embedded into ExcelMerger.exe):
node --version
echo.

echo [INFO] npm install ...
call npm install
if errorlevel 1 goto :fail

echo [INFO] npm run build  (dist\index.html) ...
call npm run build
if errorlevel 1 goto :fail
if not exist "dist\index.html" goto :fail

echo [INFO] npm run build:runner  (server\excel\excel-runner.ps1) ...
call npm run build:runner
if errorlevel 1 goto :fail

echo [INFO] npm run build:exe  (ExcelMerger.exe, takes about 30-60 seconds) ...
call npm run build:exe
if errorlevel 1 goto :fail

echo.
echo [OK] ExcelMerger.exe is ready. Double-click start.bat to run the app.
echo      To make the folder and the zip to hand out: npm run release
if /i not "%~1"=="nopause" pause
popd
endlocal
exit /b 0

:nonode
echo [ERROR] Node.js was not found. Install Node.js 24 LTS (x64) from https://nodejs.org/ first.
if /i not "%~1"=="nopause" pause
popd
endlocal
exit /b 1

:oldnode
echo [ERROR] Node.js v%NODE_MAJOR% is too old to build the exe. Install Node.js 24 LTS (x64) from https://nodejs.org/ first.
if /i not "%~1"=="nopause" pause
popd
endlocal
exit /b 1

:fail
echo.
echo [ERROR] Build failed. See the messages above.
echo         If "cannot replace ... ExcelMerger.exe": close the running app first.
if /i not "%~1"=="nopause" pause
popd
endlocal
exit /b 1

:nofolder
echo [ERROR] Cannot open the folder of build.bat: %~dp0
pause
endlocal
exit /b 1
