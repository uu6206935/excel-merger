# R5: Packaging as a Windows .exe + .bat launcher, and an app-like window

Target: "Excel統合ツール" (Excel Merger). React/Vite single-file UI + a small Node server (Node built-ins only),
started by `start.bat`, optionally shipped as `ExcelMerger.exe`.
Machine used for all tests: Windows 11 Home 10.0.26200 (Japanese, OEM/ANSI code page 932), Node v24.16.0 x64,
npm 11.13, Edge (`C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`), Chrome, Windows Terminal set as the
default terminal ("let Windows decide"), Windows PowerShell 5.1.26100, Defender real-time on, Smart App Control Off,
Windows SDK 10.0.26100 installed (so `signtool.exe` exists), Excel 16 (M365).

Everything below was run hands-on in `scratchpad\research\r5` unless marked "(docs)" or "(not tested)".

---------------------------------------------------------------------------------------------------------

## 0. TL;DR / recommendation

1. **Package with Node SEA (Single Executable Application), classic Node 24 flow**: esbuild bundles
   `server/server.mjs` -> one CommonJS file, `node --experimental-sea-config`, copy `node.exe`, `signtool remove /s`
   (optional), `postject` inject, then (optional) icon/version with `resedit`. Script: `r5/build-exe.mjs`
   (tested end-to-end through `build.bat`: 31-38 s for the exe step, exe 95.6 MB = 91.1 MiB, zip 39.1 MB). `dist/index.html` and the `.ps1`
   scripts are embedded as SEA assets and read with `node:sea` `getAsset()`.
2. **Also support a "runtime" layout** (`build-exe.mjs --layout runtime`): the *unmodified, OpenJS-signed*
   `runtime\node.exe` + `app\server.cjs`. Same code, no SEA. Use it when a PC's Smart App Control / AV blocks
   the unsigned SEA exe (the SEA exe cannot keep Node's signature). `start.bat` picks whichever is present.
3. **App window = Edge `--app=http://localhost:PORT/` with a dedicated `--user-data-dir`**
   (`%LOCALAPPDATA%\ExcelMerger\edge-profile`), fallback Chrome `--app`, fallback default browser.
   Use **`localhost`, not `127.0.0.1`** in the URL: Edge writes a Mark-of-the-Web `ZoneId=3` on every download
   from `http://127.0.0.1` (exported .xlsx would open in Excel Protected View) but none for `http://localhost`
   (tested). Keep the server bound to `127.0.0.1`.
4. **Server lifetime = window lifetime**: page holds an `EventSource('/api/events')` and sends
   `navigator.sendBeacon('/api/bye')` on `pagehide`. Server exits 4 s after the last SSE connection closed *with*
   a bye (reload reconnects within milliseconds, tested), or after 30 min if connections vanished without a bye
   (tab discarded/sleeping, crash). Do not rely on the browser PID (a 2nd window's msedge.exe exits immediately).
5. **Launcher mode** (`ExcelMerger.exe --detach`, used by start.bat): the visible console process starts a
   detached copy of itself (no console, log to `%LOCALAPPDATA%\ExcelMerger\logs\server.log`), waits for its READY
   over IPC (~170-260 ms), prints a Japanese "started" line and exits, so after double-clicking start.bat only
   the app window stays. Errors before READY are shown and the bat `pause`s. `start.bat debug` = foreground server.
6. **`.bat` files must be ASCII-only (and CRLF)**. Tested in fresh consoles: a UTF-8 .bat with `chcp 65001`
   mis-parses Japanese lines (fragments of lines are executed as commands), a BOM breaks `@echo off`, and a
   Shift-JIS .bat only works on cp932 systems and breaks when saved from a UTF-8 editor. Print Japanese from
   node/exe instead (always correct: libuv writes with WriteConsoleW, tested under cp932 and cp65001).
7. Unsigned exe + MOTW: tell users to **unblock the zip before extracting** (Properties -> "許可する"/Unblock or
   `Unblock-File`); locally built exe has no MOTW and runs without any prompt (tested). Defender scanned the
   exe clean, but SEA binaries are an AV-heuristics target (Stealit campaign 2025) - keep the runtime layout
   as plan B, and code-sign if this ever goes beyond personal use.

---------------------------------------------------------------------------------------------------------

## 1. Node SEA on Node v24.16.0 (tested)

### 1.1 What Node 24.16 supports

* `node --help` on v24.16.0 lists `--experimental-sea-config` only; **no `--build-sea`** (that flag and
  `"mainFormat": "module"` (ESM main) arrived in **v25.5.0** (docs)). `build-exe.mjs` detects `--build-sea`
  in `node --help` and uses it automatically on a newer Node; otherwise the classic flow below.
* SEA main script must be **CommonJS**; inside it `require()` only loads **built-in** modules
  (docs; `createRequire(__filename)` is the documented escape hatch, not needed when everything is bundled).
* `node:sea` API present in 24.16: `isSea, getAsset, getRawAsset, getAssetAsBlob, getAssetKeys`.
  - `getAsset(key)` -> ArrayBuffer copy (`Buffer.from(ab)`), `getAsset(key, 'utf8')` -> string,
    `getRawAsset(key)` -> no-copy ArrayBuffer (do not write to it).
* Inside SEA: `__filename === process.execPath`, `process.argv = [exe, exe, ...userArgs]` (observed).
* `useCodeCache: true` => dynamic `import()` does not work (docs). Don't use `import()` in server code.
* sea-config **relative paths resolve against the current working directory, not the config file's folder**
  (tested: running from another folder -> `Cannot read main script build/sea/server.cjs`). The build script
  writes absolute paths.
* `"execArgvExtension": "none"` works on 24.16: with `NODE_OPTIONS="--require=./does-not-exist.js"` the exe
  still started (plain `node` fails with that env). Default is `"env"` (NODE_OPTIONS honoured) - set `"none"`
  so a developer's NODE_OPTIONS on some PC cannot break/alter the exe.

### 1.2 Build flow (exact commands; `build-exe.mjs` does all of it)

```text
esbuild server/server.mjs --bundle --platform=node --format=cjs --target=node24
        --define:import.meta.url=__import_meta_url
        --banner:js="const __import_meta_url = require(\"node:url\").pathToFileURL(__filename).href;"
        --outfile=build/sea/server.cjs
node --experimental-sea-config build/sea/sea-config.json          -> build/sea/sea-prep.blob
copy "%ProgramFiles%\nodejs\node.exe" release\ExcelMerger\ExcelMerger.exe   (fs.copyFileSync(process.execPath))
signtool remove /s ExcelMerger.exe                                  (optional)
npx postject ExcelMerger.exe NODE_SEA_BLOB build/sea/sea-prep.blob ^
    --sentinel-fuse NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2
(optional) resedit: icon + VERSIONINFO   -- AFTER postject
```

JS API used for injection (avoids npx/cmd quoting):
```js
const { inject } = require('postject');           // postject@1.0.0-alpha.6 (wasm LIEF)
await inject(exePath, 'NODE_SEA_BLOB', blobBuffer, {
  sentinelFuse: 'NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2',
});
```

`sea-config.json` as generated (absolute paths in the real build; `r5/sea-config.json` is a relative-path
example that must be run with cwd = project root):
```json
{
  "main": "<root>/build/sea/server.cjs",
  "output": "<root>/build/sea/sea-prep.blob",
  "disableExperimentalSEAWarning": true,
  "useSnapshot": false,
  "useCodeCache": true,
  "execArgvExtension": "none",
  "assets": {
    "index.html": "<root>/dist/index.html",
    "excel-bridge.ps1": "<root>/server/excel-bridge.ps1"
  }
}
```

**Bundling `server.mjs` (ESM) into one CJS file** - the only gotchas:
* `import.meta.url` is emptied by esbuild in CJS output -> `fileURLToPath(undefined)` throws. Fix with the
  `define` + `banner` above (inside SEA `__filename` is the exe path, so `import.meta.url` becomes the exe's URL).
* No top-level `await` in server code (not allowed in CJS output).
* Detect SEA without importing `node:sea` statically (works in ESM dev run, in the CJS bundle, and in SEA):
  ```js
  const sea = typeof process.getBuiltinModule === 'function' ? process.getBuiltinModule('node:sea') : null;
  const IS_SEA = !!(sea && sea.isSea());
  const APP_DIR = IS_SEA ? path.dirname(process.execPath)
                         : path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  ```
* npm deps are fine if pure JS (esbuild inlines them). Native addons (`.node`, e.g. `winax`) cannot be required
  from SEA directly (docs: embed as asset, write to temp, `process.dlopen`) - avoid them.

### 1.3 Signature handling (tested)
* Official `node.exe` is Authenticode-signed (OpenJS Foundation, `Valid`). Any injection invalidates it.
* With `signtool remove /s` first: result `NotSigned`, runs fine.
* **Without signtool** (`--keep-signature`): postject prints `warning: The signature seems corrupted!`, LIEF
  drops the certificate anyway -> `NotSigned`, same size, runs fine. So signtool/Windows SDK is **not required**.
* signtool location if wanted: `C:\Program Files (x86)\Windows Kits\10\bin\<ver>\x64\signtool.exe`
  (build script searches `where signtool` then Windows Kits).

### 1.4 Icon / version info (tested)
* `resedit@3.1.0` + `pe-library@2.0.1` (pure JS, **ESM-only** -> load with `await import()` from the build
  script). Replaces the Node icon group and VERSIONINFO (FileDescription/ProductName "Excel統合ツール",
  1.0.0.0) - Explorer shows the new icon and description.
* **Order matters**: resedit *before* postject -> postject prints
  `error: Relocation corrupted: BlockSize is out of bound the binary's virtual size` (exe still ran, but don't
  risk it). resedit *after* postject -> clean, SEA assets intact (verified `/api/info` assetKeys, 3 MB index.html
  served, ps1 ran). The NODE_SEA_BLOB lives in an RT_RCDATA resource and resedit preserves it.
* `r5/make-ico.mjs` writes a PNG-in-ICO (16/32/48/256) with no dependencies if you need a quick placeholder.

### 1.5 Numbers (this PC)

| item | value |
|---|---|
| node.exe v24.16.0 | 92,279,112 B |
| SEA exe (3.1 MB index.html + 1.5 KB ps1 embedded) | 95,572,480 B (95.6 MB = 91.1 MiB) |
| release zip (exe + dist/index.html copy + start.bat) | 39,107,350 B (39.1 MB) (test HTML is random base64 = incompressible; a real Vite bundle compresses better) |
| build time (bundle + blob + copy + signtool + postject + resedit + smoke test) | 31-38 s (postject alone 15-27 s) |
| build.bat end to end (npm install + UI build + exe + zip) | 44-48 s |
| startup to "READY" (listening), warm, median | SEA 99-172 ms (varies with machine load); `node server.cjs` 86-156 ms |
| startup, first run of a freshly built/copied exe | 1.7-2.6 s (Defender scans the new 95 MB file once) |
| useCodeCache true vs false | ~99-154 ms vs ~131-172 ms (small win; keep true) |
| serving the 3 MB embedded index.html | 6 ms (curl, local) |
| `--detach` launcher: start -> READY + Edge spawned | 166-258 ms |
| Edge `--app` window -> page connected (SSE) | ~1 s (fresh profile ~16 MB on first launch) |
| powershell.exe run of the embedded ps1 | 300-470 ms; with Excel COM create/quit ~2.5 s |

### 1.6 Verified behaviours
* Exe runs from `...\jp path\テスト フォルダ (日本語)\Excel統合ツール v1\ExcelMerger.exe` (Japanese + spaces +
  parentheses): `APP_DIR`, assets, ps1 extraction to a Japanese folder (`データ フォルダ\ps\...`), PowerShell run
  and Japanese JSON round trip all OK.
* Embedded ps1 -> written to `%LOCALAPPDATA%\ExcelMerger\ps\<sha256-12>-excel-bridge.ps1` once (content hash
  in the name => versions never clash; write to `.tmp` then rename), bytes copied as-is so the UTF-8 BOM stays.
* PowerShell spawned from the SEA exe created its own Excel instance via COM, read `Version=16.0`, got its PID via
  `GetWindowThreadProcessId(Hwnd)`, quit it (needed a force-kill of *that PID only* because Excel was still alive
  500 ms after `Quit()` - wait longer, ~5-10 s, before killing).
* Defender `MpCmdRun -Scan -ScanType 3 -File ExcelMerger.exe` -> "found no threats".
* Locally built exe has no `Zone.Identifier` stream -> no SmartScreen prompt when run on the build PC.
* `--gui-subsystem` (patch PE `Subsystem` word at PE+0x5C to 2): exe runs with no console at all; `console.log`
  to missing handles does not crash; piped stdout still works; PowerShell child still runs hidden. Works, but
  errors become invisible -> prefer the `--detach` launcher (same UX, errors visible). Kept as an option.

---------------------------------------------------------------------------------------------------------

## 2. Alternatives (compared)

| option | tested | exe size | build | startup | verdict |
|---|---|---|---|---|---|
| **Node SEA** (classic flow, Node 24) | yes | 91.1 MiB | 31-38 s | ~100-170 ms | **use** - official, no third-party binaries, same Node as dev |
| **"runtime" layout**: signed `runtime\node.exe` + `app\server.cjs` | yes | 88.0 MiB node.exe + files | 2.4 s | same as node | **use as plan B** - keeps OpenJS signature (SAC/AV friendly); not a single file, Task Manager shows "Node.js" |
| @yao-pkg/pkg 6.23.0 (target node22-win-x64) | yes | 57.9 MiB | 23 s (first run downloads a 57 MB patched Node from github.com/yao-pkg/pkg-fetch) | 169 ms median | reference-only - works (assets via `/snapshot` fs, `pkg.assets`), smaller exe with node22, but third-party patched Node binaries, Node 22 EOL 2027-04; node24 base is 91 MB anyway |
| bun 1.4.2 `bun build --compile` (portable, from npm into scratch) | yes | 82.1 MiB | 1.1 s | 99 ms | reject - fast, http/SSE/child_process worked, but a different runtime (reports node v26.3.0 compat) and asset embedding needs Bun-only `import x from './f' with {type:'file'}`/`Bun.file`; `node:sea` absent -> code divergence |
| nexe 5.0.0-beta.4 | no | - | - | - | reject - last GitHub release v3.3.3 (2017), npm beta 2025-03; prebuilt Node binaries only via third-party `urbdyn/nexe_builds` (0.4.0, 2025-03), otherwise compiles Node from source (VS Build Tools + Python, ~1 h) |
| deno compile | no (not installed) | ~70-90 MB (docs/experience) | - | - | reject - Node server would have to run under Deno's npm/node compat + permission flags; no benefit over SEA |
| Node 25.5+ `node --build-sea` | no (not on this PC) | same as SEA | faster (no postject) | same | future: `build-exe.mjs` switches to it automatically when `node --help` lists `--build-sea` |

---------------------------------------------------------------------------------------------------------

## 3. App-like window (tested)

### 3.1 Finding the browser
Order used by `findChromium()` in `r5/app/server/server.mjs`:
1. `reg query "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe" /ve`
2. same under `HKLM\SOFTWARE\...` and `HKLM\SOFTWARE\WOW6432Node\...`
   (this PC: HKLM -> `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`; HKCU none)
3. `%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe`, `%ProgramFiles%\...`, `%LOCALAPPDATA%\...`
4. Same for Chrome (`chrome.exe`, `Google\Chrome\Application`; this PC: HKLM -> `C:\Program Files\Google\Chrome\Application\chrome.exe`)
5. Fallback: `cmd /d /s /c start "" "URL"` (default browser, normal tab; `windowsVerbatimArguments: true`).
   (`reg.exe` output is parsed with `/REG_(?:EXPAND_)?SZ\s+(.+)$/m`; expand `%VAR%`.)

### 3.2 Launch
```js
spawn(msedge, [
  `--app=${url}`,                                   // http://localhost:PORT/
  `--user-data-dir=${LOCALAPPDATA}\\ExcelMerger\\edge-profile`,
  '--no-first-run', '--no-default-browser-check', '--disable-sync',
  '--window-size=1400,900',
], { stdio: 'ignore', detached: true }).unref();
```
Observed: one window, class `Chrome_WidgetWin_1`, title = page `<title>` ("Excel統合ツール"), favicon from the page,
no address bar, 1400x900 at (10,10), no first-run/sign-in window with a fresh profile. Chrome with its own
`--user-data-dir` behaves the same and does **not** touch the user's running Chrome (tested while the user's Chrome
had 63 processes).

**Why a dedicated profile**: the launched msedge.exe *is* the browser process (it exited 0 when the window closed -
no background mode), `--window-size` is honoured, the user's extensions/tabs/sessions are untouched, localStorage/
IndexedDB of the app persist in the profile. Cost: 16 MB (Edge) / 57 MB (Chrome) profile folder on first launch.
**But do not use the PID for lifetime**: a second `msedge --app` with the same profile hands the URL to the running
browser and exits immediately (observed `pid 3456 exited (code 0)` while its window stayed open). Enterprise policy
`UserDataDir` could also override the flag. Hence the SSE design below.

### 3.3 `--app` with `file://` (standalone mode, tested)
| `--app=` value (Japanese path with spaces) | result |
|---|---|
| raw path `C:\...\テスト フォルダ (日本語)\...\index.html` | plain "新しいタブ - Microsoft Edge" window (not loaded) |
| `file:///C:/.../テスト フォルダ (日本語)/.../index.html` (unencoded) | plain "新しいタブ" window |
| `([Uri]$path).AbsoluteUri` (percent-encoded) | **app window, page loaded** |

The bat gets the encoded URL from PowerShell (ASCII output, safe in `for /f`):
```bat
for /f "usebackq delims=" %%u in (`powershell -NoProfile -NonInteractive -Command "([Uri](Resolve-Path -LiteralPath 'dist\index.html').ProviderPath).AbsoluteUri" 2^>nul`) do set "APPURL=%%u"
start "" "%EDGE%" --app="%APPURL%" --user-data-dir="%EM_DATA_DIR%\edge-profile" --no-first-run --no-default-browser-check --window-size=1400,900
```
Verified via start.bat from the Japanese folder (window title "Excel統合ツール"). Last resort: `start "" "%~dp0dist\index.html"`.

### 3.4 localhost vs 127.0.0.1 and Mark-of-the-Web on exported files (tested)
Page in an Edge app window downloads (a) a Blob via `<a download>` and (b) a server response with
`Content-Disposition: attachment`:
* origin `http://127.0.0.1:5220` -> both files get `Zone.Identifier: ZoneId=3, HostUrl=http://127.0.0.1:5220/...`
  => Excel opens them in **Protected View** (保護ビュー), macros in xlsm blocked.
* origin `http://localhost:5220` -> **no Zone.Identifier stream** on either file.
=> open the UI at `http://localhost:<port>/` (server still listens on 127.0.0.1 only; Chromium resolves localhost
to loopback itself - page loaded fine). Also keep the Host-header allow-list (localhost/127.0.0.1/[::1]) from the
reference server. Server-side saves (Node `fs.writeFile`, like the reference's `POST /api/output`) never get MOTW.
(Not tested: File System Access API `showSaveFilePicker` writes - Chromium may quarantine those too; test in R-UI.)

### 3.5 Auto-shutdown when the window closes (tested)
Heartbeat timers alone are fragile: Chromium throttles timers in hidden/minimised pages (intensive throttling
aligns chained timers to 1/min after 5 min hidden), and Edge "sleeping tabs"/Chrome Memory Saver discard tabs
without unload events. Design used (all Node built-ins):

Server (`r5/app/server/server.mjs`):
```js
const IDLE_EXIT_MS = 30 * 60 * 1000, EXIT_GRACE_MS = 4000;
const clients = new Set(); let byeAt = 0, exitTimer = null;
function scheduleExitIfIdle(reason) {
  if (clients.size > 0) return;
  clearTimeout(exitTimer);
  const wait = Date.now() - byeAt < 15000 ? EXIT_GRACE_MS : IDLE_EXIT_MS;   // bye seen => real close
  exitTimer = setTimeout(() => { if (clients.size === 0) process.exit(0); }, wait);
}
// GET /api/events
res.writeHead(200, { 'Content-Type': 'text/event-stream; charset=utf-8', 'Cache-Control': 'no-store', Connection: 'keep-alive' });
res.write('retry: 1000\n\n'); clients.add(res); clearTimeout(exitTimer);
const ping = setInterval(() => res.write(': ping\n\n'), 20000);
req.on('close', () => { clearInterval(ping); clients.delete(res); scheduleExitIfIdle('closed'); });
// POST /api/bye  -> byeAt = Date.now(); res.writeHead(204).end(); scheduleExitIfIdle('bye');
```
Page:
```js
const es = new EventSource('/api/events');
es.onerror = () => showBanner('サーバーとの接続が切れました');     // server gone / restarted
addEventListener('pagehide', (e) => { if (!e.persisted) navigator.sendBeacon('/api/bye', ''); });
```
Observed logs:
* reload (F5 simulated): `disconnected` -> `bye` -> `connected` within the same second -> no exit.
* close window (WM_CLOSE): `bye` + `disconnected` -> Edge process exits -> server `exiting` 4 s later.
* 2 windows: closing one keeps the server; closing both exits it.
* SSE via curl killed without bye -> 30-min timer (not 4 s), reconnect cancels it.
Order of `bye` vs socket close varies; the code handles both. Node's default `server.requestTimeout`
(300 s) does not cut long SSE responses (it applies to receiving the request).

### 3.6 Single instance (tested)
On `EADDRINUSE`, `GET http://127.0.0.1:<port>/api/ping`; if it answers `{app:'ExcelMerger'}` just open another app
window to it and exit; else try port+1 (foreign process, e.g. the user's interview-scheduler on 5173 - it was
running during the tests, so **do not use 5173** as the base port). Keep the port stable because localStorage/
IndexedDB are per origin (`http://localhost:5180` != `http://localhost:5181`).

### 3.7 Launcher mode `--detach` (tested with the SEA exe, runtime node.exe and system node)
```js
function launcherMain() {                        // runs when --detach and EM_CHILD !== '1'
  const fd = fs.openSync(path.join(DATA_DIR, 'logs', 'server.log'), 'a');   // rotate at 1 MB
  const child = spawn(process.execPath, IS_SEA ? [] : [process.argv[1]], {
    detached: true,                               // DETACHED_PROCESS|CREATE_NEW_PROCESS_GROUP: no console
    windowsHide: true, stdio: ['ignore', fd, fd, 'ipc'], env: { ...process.env, EM_CHILD: '1' },
  });
  child.on('message', (m) => { if (m.ready) { console.log(`起動しました: ${m.url}`); child.unref(); setTimeout(() => process.exit(0), 1500); } });
  child.on('exit', (code) => fail(`exit code ${code}`));   // ignore after success (flag)
}
// in the server, after listen():  if (process.send) { process.send({ ready: true, url }); process.disconnect(); }
// and in the "already running" branch: process.send({ ready: true, url, window: 'reused running instance' })
```
* IPC works inside SEA (spawning `process.execPath` with no script argument re-runs the embedded app).
* Child spawns PowerShell with pipes + `windowsHide:true` -> libuv adds `CREATE_NO_WINDOW` -> no console flash.
* Console text (Japanese) captured from the real console buffer: `起動しました / started: http://127.0.0.1:5210/ (window: edge, 187 ms)`.

### 3.8 Console hosting on Windows 11 (tested)
* A new console (double-click on a .bat) is hosted by **Windows Terminal** here (window class
  `CASCADIA_HOSTING_WINDOW_CLASS`, reused an existing WindowsTerminal.exe). `start "" /min prog.exe` instead gave a
  **minimised classic conhost** window (`ConsoleWindowClass`, iconic). With `--detach` neither matters.

---------------------------------------------------------------------------------------------------------

## 4. `.bat` launcher (tested)

### 4.1 Encoding test matrix (fresh console = what a double-click gets; buffer read back with
`AttachConsole` + `ReadConsoleOutputCharacterW`, `r5/battest/console-read.ps1`)

| file | conhost | Windows Terminal (default) |
|---|---|---|
| UTF-8 (no BOM) + `chcp 65001` | Japanese echo OK **but** line after chcp split: `'�名・学籍番号）"' is not recognized...`, a later `rem` line fragment `'うえおかきくけこ' is not recognized` was **executed** | same |
| UTF-8 **with BOM** + `chcp 65001` | `@echo off` ignored (all commands echoed) + same splits | same |
| UTF-8, no chcp | mojibake (`譌･譛ｬ隱槭...`) + splits | same |
| Shift-JIS (cp932), no chcp | **correct** (only on cp932 systems) | correct |
| ASCII bat running `node jp.js` (console.log Japanese), with or without chcp 65001 | **correct** | correct |
| ASCII bat `node jp.js | findstr` (piped) | garbled under cp932, correct under chcp 65001 | same |

(Redirected `cmd /c x.bat > file` tests are misleading: they inherit the caller's code page - the tool's shell
was already 65001.) Conclusion: **ASCII-only .bat, CRLF line endings, Japanese only from node/exe**; `chcp 65001`
is harmless in an ASCII bat and helps only if node output is piped. Note `%~dp0` with Japanese folders works in
ASCII bats regardless of code page (tested). Add `*.bat text eol=crlf` to `.gitattributes` (LF-only bats have a known
label-lookup bug; not reproduced in a small test here). Other rules kept from the reference: `pushd "%~dp0"`
(UNC-safe), no `echo` inside `( ... )` blocks, `goto` style error handling, `where node` + major-version check.

### 4.2 Deliverables
* `r5/launcher/start.bat` (ASCII, 109 lines): `ExcelMerger.exe --detach` -> `runtime\node.exe app\server.cjs --detach`
  -> installed `node server\server.mjs --detach` (Node >= 20) -> standalone `dist\index.html` in Edge `--app`
  (percent-encoded file URL) -> default browser. `start.bat debug` = `--foreground`. Sets `EM_FROM_BAT=1` so the exe
  does not wait for Enter itself (the bat `pause`s on errorlevel 1). Tested from a Japanese folder in all four modes
  (exe, runtime, node, standalone) incl. window close -> server exit.
* `r5/launcher/build.bat` (= `r5/sample-project/build.bat`): Node >= 22 check, `npm install`, `npm run build`,
  `node tools\build-exe.mjs --icon assets\app.ico`, `Compress-Archive` -> `release\ExcelMerger.zip`; `build.bat nopause`
  for scripts. Tested end-to-end (exit 0, 44-48 s).
* Release folder produced by `build-exe.mjs` (SEA layout):
  ```
  release\ExcelMerger\ExcelMerger.exe   (server + UI + ps1 embedded)
  release\ExcelMerger\start.bat         (copied; build fails if it is not ASCII)
  release\ExcelMerger\README.txt
  release\ExcelMerger\dist\index.html   (standalone fallback)
  ```
  runtime layout (`--layout runtime`): `runtime\node.exe`, `app\server.cjs`, `server\*.ps1`, `dist\index.html`,
  `start.bat`, `README.txt` (node.exe signature stays `Valid`, tested).

---------------------------------------------------------------------------------------------------------

## 5. SmartScreen / Defender / Smart App Control / MOTW

* **Locally built** exe: no `Zone.Identifier` -> SmartScreen never asks on the build PC (tested). Defender scanned it
  clean; the *first* launch of every new/copied exe costs 1.7-2.6 s (real-time scan), later launches ~0.1 s.
* **Copied to other PCs via download / Teams / mail / browser**: the zip gets MOTW; Explorer's built-in "Extract all"
  propagates it to the extracted files -> unsigned `ExcelMerger.exe` shows SmartScreen "Windows によって PC が保護されました"
  (詳細情報 -> 実行) and the .bat may show the "Open File - Security Warning" dialog. Fix: **unblock the zip before
  extracting** (right-click -> Properties -> 許可する / `Unblock-File .\ExcelMerger.zip`), or copy via a USB stick / file
  share. Put this in README.txt.
* **Smart App Control** (Windows 11 clean installs; Off on this PC): blocks unsigned executables without reputation at
  launch (docs/Microsoft FAQ + GitHub reports). An SEA exe is unsigned by construction => blocked when SAC is On. The
  runtime layout (unmodified OpenJS-signed node.exe) is the workaround; alternatively sign the exe
  (`signtool sign /fd SHA256 /a ...` needs a code-signing certificate).
* **AV heuristics**: SEA exes are "modified node.exe + appended data"; a 2025 malware campaign (Stealit) used Node SEA,
  and a 2026 issue reports a scanner flagging SEA PE layout ("export table overlaps with executable headers"). Defender
  was fine here. If a corporate AV quarantines the exe: use the runtime layout.
* Behavioural pattern to keep low-profile: an unsigned exe spawning hidden `powershell.exe -ExecutionPolicy Bypass`
  resembles malware. Prefer the script-block invocation below (no `-ExecutionPolicy Bypass`, no `-EncodedCommand`),
  keep scripts as readable files under `%LOCALAPPDATA%\ExcelMerger\ps`, and log what runs.

---------------------------------------------------------------------------------------------------------

## 6. PowerShell from the exe (tested notes for the Excel-engine topic)

* Use the absolute path `%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe` (PATH may be odd).
* Return data through a **UTF-8 JSON file** (`[IO.File]::WriteAllText($OutJson, $json, (New-Object Text.UTF8Encoding $false))`);
  set `[Console]::OutputEncoding = New-Object Text.UTF8Encoding $false` in the script so stdout is UTF-8 too.
  PowerShell's own early errors on stderr are in cp932 -> decode with `new TextDecoder('shift_jis')` (available in the
  official full-ICU Node/SEA build, tested) when the bytes are not valid UTF-8.
* **Execution policy**: a GPO/"Restricted" policy blocks `-File x.ps1` even with `-ExecutionPolicy Bypass` when set by
  Group Policy. Script blocks built from text are not subject to execution policy - tested under
  `-ExecutionPolicy Restricted`: `-File` failed, this worked (and `$PSCommandPath` is empty in this mode):
  ```js
  const q = (s) => `'${s.replace(/'/g, "''")}'`;
  const cmd = `$ErrorActionPreference='Stop'; & ([ScriptBlock]::Create([IO.File]::ReadAllText(${q(ps1)}, [Text.Encoding]::UTF8))) -OutJson ${q(outJson)}`;
  spawn(POWERSHELL, ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', cmd], { windowsHide: true });
  ```
  (Not testable here: AppLocker/WDAC Constrained Language Mode blocks `New-Object -ComObject` entirely - the Excel
  engine must degrade to the JS engine with a clear message.)
* Excel COM: `New-Object -ComObject Excel.Application` (own instance), `Visible=$false`, `DisplayAlerts=$false`, PID via
  `GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$pid)`, `Quit()` + `FinalReleaseComObject` + `GC` in `finally`, then
  wait up to ~10 s and kill only that PID. One PowerShell process per export job (startup ~0.3-0.5 s each).

---------------------------------------------------------------------------------------------------------

## 7. Pitfalls checklist

1. SEA on Node 24 = CommonJS only; bundle with esbuild; define `import.meta.url` (banner) or APP_DIR detection breaks.
2. sea-config relative paths are relative to the **cwd**; write absolute paths.
3. Apply icon/version (resedit) **after** postject; resedit/pe-library are ESM-only (`await import`).
4. postject takes 15-27 s (wasm); it warns about the signature if signtool isn't used - harmless.
5. `fs.rmSync(OUT)` fails with EBUSY while the old exe runs - close the app before rebuilding (message in script).
6. The embedded Node is whatever `process.execPath` the build runs with (build with x64 Node 24 LTS; ARM64 PCs run the
   x64 exe under emulation).
7. Set `execArgvExtension: "none"` (NODE_OPTIONS on a user PC otherwise changes the exe; verified).
8. Don't use dynamic `import()` in server code with `useCodeCache: true`.
9. First run of a new exe: +1.7-2.6 s Defender scan - show a splash/"starting" message, don't time out < 30 s.
10. Use `http://localhost:<port>` for the UI (127.0.0.1 => MOTW ZoneId=3 on downloads => Protected View).
11. Edge `--app` with `file://` needs a percent-encoded URL (`[Uri]::AbsoluteUri`).
12. Browser PID is not the window lifetime (2nd window's msedge.exe exits at once); use SSE + pagehide beacon.
13. Timers in hidden windows are throttled; never auto-exit on a short heartbeat timeout.
14. Don't reuse port 5173 (the reference app runs there); keep the port fixed for a stable localStorage origin.
15. `.bat`: ASCII-only, CRLF, no echo inside `( )`, `pushd "%~dp0"`; Japanese output from node/exe only.
16. No BOM in .bat; ps1 needs BOM when it contains Japanese (build script rejects BOM-less non-ASCII ps1).
17. Double-clicking a file inside a zip (Explorer's zip view) runs it from a temp copy without `dist\` - README must say "unzip first".
18. Kill processes only by PID you created (tests used `win-inspect.ps1 -Match <our profile dir>` to find/close only our windows).

---------------------------------------------------------------------------------------------------------

## 8. Files in `scratchpad\research\r5`

| path | what |
|---|---|
| `build-exe.mjs` | **the build script** (SEA classic flow / `--build-sea` / `--layout runtime`, icon, GUI subsystem, smoke test, release folder) |
| `sea-config.json` | example config (relative paths => run with cwd = project root); `build-sea-config.generated.json` = real generated one |
| `launcher/start.bat`, `launcher/build.bat` | final ASCII launchers |
| `app/server/server.mjs` | prototype server: SEA detection, asset serving, ps1 extraction + run, Edge/Chrome/default-browser app window, single instance, SSE+bye auto-exit, `--detach` launcher |
| `app/server/excel-bridge.ps1` | embedded ps1 (UTF-8 BOM, JSON-file result, optional Excel COM check with own-PID cleanup) |
| `app/gen-index-html.mjs` | makes a 3 MB single-file-like `dist/index.html` with the SSE/pagehide page code |
| `sample-project/` | complete mini project: `package.json`, `server/`, `tools/build-exe.mjs`, `assets/app.ico`, `start.bat`, `build.bat`, `release/ExcelMerger/*` + `release/ExcelMerger.zip`, `release-runtime/ExcelMerger/*` |
| `out/ExcelMerger.exe` | built SEA exe (icon, version info); `out/gui/`, `out/keepsig/`, `out/nocache/` variants |
| `alt/out-pkg/ExcelMerger-pkg22.exe`, `alt/out-bun/ExcelMerger-bun.exe` | alternative builds used for the comparison |
| `measure-startup.mjs` | spawn N times, time to READY |
| `win-inspect.ps1` | list/capture/close only windows of browsers started with a given profile dir |
| `battest/console-read.ps1` | run a .bat in a fresh console (conhost or default terminal) and read the screen buffer back |
| `battest/*.bat`, `battest/console_*.txt` | encoding test files and captured results |
| `ps-scriptblock-test.mjs` | execution-policy-proof PowerShell invocation test |
| `motw/dl-server.mjs` | MOTW-on-download test (127.0.0.1 vs localhost) |
| `make-ico.mjs`, `app.ico` | dependency-free ICO generator + icon |

---------------------------------------------------------------------------------------------------------

## 9. Sources
* Node.js v24 docs, Single executable applications: https://nodejs.org/docs/latest-v24.x/api/single-executable-applications.html
* Node.js latest docs (`--build-sea`, `mainFormat` since v25.5.0): https://nodejs.org/docs/latest/api/single-executable-applications.html
* postject: https://github.com/nodejs/postject (npm 1.0.0-alpha.6)
* resedit-js: https://github.com/jet2jet/resedit-js (npm resedit 3.1.0, pe-library 2.0.1)
* @yao-pkg/pkg: https://github.com/yao-pkg/pkg (6.23.0), binaries https://github.com/yao-pkg/pkg-fetch/releases (v3.6)
* nexe: https://github.com/nexe/nexe ; prebuilt https://github.com/urbdyn/nexe_builds
* Bun single-file executables: https://bun.sh/docs/bundler/executables
* Smart App Control FAQ: https://support.microsoft.com/en-us/windows/security/threat-malware-protection/smart-app-control-frequently-asked-questions
* SEA PE scan issue: https://github.com/nodejs/single-executable/issues/127
* Stealit campaign abusing Node SEA: https://www.fortinet.com/blog/threat-research/stealit-campaign-abuses-nodejs-single-executable-application
