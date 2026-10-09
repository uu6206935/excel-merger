@echo off
rem ===========================================================================
rem  Excel Merger (Excel Tougou Tool) - launcher.  Double-click this file.
rem
rem  KEEP THIS FILE ASCII-ONLY WITH CRLF LINE ENDINGS. cmd.exe mis-parses UTF-8
rem  lines with Japanese text (parts of a line are run as commands), a UTF-8 BOM
rem  breaks "@echo off", and a Shift-JIS copy breaks when saved from a UTF-8
rem  editor. Japanese messages are printed by ExcelMerger.exe / node instead.
rem
rem  Order: ExcelMerger.exe (single exe)
rem         -> runtime\node.exe + app\server.cjs ("runtime" layout)
rem         -> installed Node.js 18+ + server\server.mjs (builds the UI first if needed)
rem         -> standalone: dist\index.html in an Edge app window (no Excel-app engine)
rem  A method that cannot be started at all (the exe blocked by Smart App Control /
rem  antivirus software, a broken node.exe ...) hands over to the next one by itself. A process that
rem  was killed or crashed exits with a NEGATIVE code (-1073741515 ...): "if errorlevel 1" would miss it,
rem  so the launch checks compare %ERRORLEVEL% with 0.
rem  "start.bat debug" keeps the server in this window so its log is visible (no hand-over).
rem ===========================================================================
setlocal
chcp 65001 >nul
rem pushd (not cd) also works for a folder on a network share (\\server\share\...).
pushd "%~dp0" || goto :nofolder
title Excel Merger
set "EM_FROM_BAT=1"
set "EM_RUN=--detach"
if /i "%~1"=="debug" set "EM_RUN=--foreground"
if "%EM_DATA_DIR%"=="" set "EM_DATA_DIR=%LOCALAPPDATA%\ExcelMerger"
set "EM_SKIP_EXE="
set "EM_SKIP_RUNTIME="
set "EM_SKIP_NODE="

:choose
if exist "ExcelMerger.exe" if not defined EM_SKIP_EXE goto :exe
if not exist "runtime\node.exe" goto :checknode
if not exist "app\server.cjs" goto :checknode
if not defined EM_SKIP_RUNTIME goto :runtime
:checknode
if defined EM_SKIP_NODE goto :standalone
goto :node

:exe
"%~dp0ExcelMerger.exe" %EM_RUN%
if not "%ERRORLEVEL%"=="0" goto :exefailed
goto :done

:exefailed
if /i "%~1"=="debug" goto :failed
echo [WARN] ExcelMerger.exe could not be started. Windows Smart App Control or antivirus software may have blocked it.
echo [WARN] Trying another way ...
set "EM_SKIP_EXE=1"
goto :choose

:runtime
rem "runtime" layout: the unmodified, OpenJS-signed node.exe + the bundled server.
"%~dp0runtime\node.exe" "%~dp0app\server.cjs" %EM_RUN%
if not "%ERRORLEVEL%"=="0" goto :runtimefailed
goto :done

:runtimefailed
if /i "%~1"=="debug" goto :failed
echo [WARN] The bundled node.exe could not be started. Trying another way ...
set "EM_SKIP_RUNTIME=1"
goto :choose

:node
if not exist "server\server.mjs" goto :standalone
where node >nul 2>nul
if errorlevel 1 goto :standalone
set "NODE_MAJOR="
for /f "tokens=1 delims=v." %%v in ('node --version 2^>nul') do if not defined NODE_MAJOR set "NODE_MAJOR=%%v"
if not defined NODE_MAJOR goto :standalone
if %NODE_MAJOR% LSS 18 goto :oldnode
if exist "dist\index.html" goto :runnode
if not exist "package.json" goto :runnode
echo [INFO] dist\index.html not found: building the app first (npm install / npm run build) ...
where npm >nul 2>nul
if errorlevel 1 goto :nonpm
if exist "node_modules\vite" goto :npmbuild
call npm install
if errorlevel 1 goto :buildfailed

:npmbuild
call npm run build
if errorlevel 1 goto :buildfailed
if not exist "dist\index.html" goto :buildfailed

:runnode
node "server\server.mjs" %EM_RUN%
if not "%ERRORLEVEL%"=="0" goto :nodefailed
goto :done

:nodefailed
if /i "%~1"=="debug" goto :failed
echo [WARN] The server could not be started with the installed Node.js. Trying standalone mode ...
set "EM_SKIP_NODE=1"
goto :choose

:oldnode
echo [WARN] Node.js v%NODE_MAJOR% is too old (18 or newer is needed). Using standalone mode.
goto :standalone

:nonpm
echo [WARN] npm was not found, so the app cannot be built here.
goto :standalone

:buildfailed
echo [WARN] The build failed (see above). Trying standalone mode.
goto :standalone

:standalone
if not exist "dist\index.html" goto :nodist
echo [INFO] ExcelMerger.exe / Node.js not found or not usable: opening dist\index.html directly (standalone mode).
echo [INFO] Files are chosen and saved by the browser; the Excel-app export engine is not available.
rem Edge --app needs a percent-encoded file:/// URL (a raw path opens a plain "New tab").
set "APPURL="
for /f "usebackq delims=" %%u in (`powershell -NoProfile -NonInteractive -Command "([Uri](Resolve-Path -LiteralPath 'dist\index.html').ProviderPath).AbsoluteUri" 2^>nul`) do set "APPURL=%%u"
set "EDGE="
for %%e in ("%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" "%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe") do if not defined EDGE if exist "%%~e" set "EDGE=%%~e"
if not defined APPURL goto :defaultbrowser
if not defined EDGE goto :defaultbrowser
start "" "%EDGE%" --app="%APPURL%" --user-data-dir="%EM_DATA_DIR%\edge-profile" --no-first-run --no-default-browser-check --window-size=1400,900
goto :done

:defaultbrowser
start "" "%~dp0dist\index.html"
goto :done

:done
popd
endlocal
exit /b 0

rem ---------------------------------------------------------------------------
rem Errors. Keep every echo outside ( ... ) blocks: an unescaped ")" in a block
rem ends the block and aborts the batch with a syntax error.
rem ---------------------------------------------------------------------------
:failed
echo.
echo [ERROR] Excel Merger could not start. See the messages above.
echo         Log folder: %EM_DATA_DIR%\logs
echo         Try "start.bat debug" to see the server log in this window.
pause
popd
endlocal
exit /b 1

:nodist
echo [ERROR] ExcelMerger.exe, Node.js and dist\index.html are all missing or not usable.
echo         Copy the whole folder again (unzip it first, do not run it from inside the zip).
echo         If Windows or antivirus software blocked ExcelMerger.exe, ask for the "runtime" version of the tool.
pause
popd
endlocal
exit /b 1

:nofolder
echo [ERROR] Cannot open the folder of start.bat:
echo         %~dp0
echo         Copy the whole folder to a local disk (for example the Desktop) and run start.bat there.
pause
endlocal
exit /b 1
