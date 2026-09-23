# Build the release: PyInstaller (onedir) -> self-test -> Inno Setup installer (if installed).
#
# Usage (from the project root, PowerShell):
#     powershell -ExecutionPolicy Bypass -File packaging\build.ps1
#     powershell -ExecutionPolicy Bypass -File packaging\build.ps1 -SkipInstaller
#
# Output:
#     dist\CUPL-Undergraduate-Thesis-Auto-Formatter\             the app folder (exe + _internal)
#     dist\installer\"CUPL-UTAF V<version>.exe"   (only if Inno Setup is found)
#
# Messages are ASCII on purpose: Windows PowerShell 5.1 reads BOM-less UTF-8 as ANSI.

param([switch]$SkipInstaller,
      [switch]$InstallerOnly)   # reuse the existing dist\CUPL-Undergraduate-Thesis-Auto-Formatter (skip steps 2-4)

# "Continue", not "Stop": Windows PowerShell 5.1 turns anything a native tool writes to stderr
# (PyInstaller logs its INFO lines there) into a terminating error. Every native call below
# is checked through $LASTEXITCODE instead.
$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

function Step($msg) { Write-Host "`n=== $msg" -ForegroundColor Cyan }
function Fail($msg) { Write-Host "FAILED: $msg" -ForegroundColor Red; exit 1 }

# ---- 1. dedicated build environment (kept separate from the dev .venv; only runtime deps + PyInstaller)
Step "Build environment"
$venv = Join-Path $PSScriptRoot ".venv"
$py = Join-Path $venv "Scripts\python.exe"
if (-not (Test-Path $py)) {
    $base = Join-Path $root ".venv\Scripts\python.exe"
    if (-not (Test-Path $base)) { Fail "project .venv not found (needed only to locate a Python 3.13 interpreter)" }
    & $base -m venv $venv
    if ($LASTEXITCODE -ne 0) { Fail "could not create the build venv" }
}
# flaky network: several retries, generous timeout; wheels are cached after the first success
& $py -m pip install --retries 8 --timeout 60 --disable-pip-version-check -q -r (Join-Path $root "requirements-build.txt")
if ($LASTEXITCODE -ne 0) { Fail "pip install failed - check the network and re-run (already downloaded wheels are cached)" }

$dist = Join-Path $root "dist"
$work = Join-Path $root "build\pyinstaller"
if (-not $InstallerOnly) {

# ---- 2. refresh the SHA-256 list of the official files that ship with the app
Step "Official-file hashes"
& $py (Join-Path $root "tools\update_integrity.py")
if ($LASTEXITCODE -ne 0) { Fail "update_integrity.py failed" }

# ---- 3. PyInstaller
Step "PyInstaller"
& $py -m PyInstaller --noconfirm --clean --distpath $dist --workpath $work (Join-Path $PSScriptRoot "CUPL-Undergraduate-Thesis-Auto-Formatter.spec")
if ($LASTEXITCODE -ne 0) { Fail "PyInstaller failed" }
$exe = Join-Path $dist "CUPL-Undergraduate-Thesis-Auto-Formatter\CUPL-Undergraduate-Thesis-Auto-Formatter.exe"
if (-not (Test-Path $exe)) { Fail "exe was not produced" }

# ---- 4. self-test of the frozen build (no window is shown)
Step "Self-test of the packaged exe"
$report = Join-Path $work "selftest.txt"
if (Test-Path $report) { Remove-Item $report -Force }
$p = Start-Process -FilePath $exe -ArgumentList @("--selftest", "`"$report`"") -PassThru -Wait
if (-not (Test-Path $report)) { Fail "self-test wrote no report (exit code $($p.ExitCode))" }
Get-Content $report -Encoding UTF8
if ($p.ExitCode -ne 0) { Fail "self-test reported failures (see above)" }

$size = (Get-ChildItem (Join-Path $dist "CUPL-Undergraduate-Thesis-Auto-Formatter") -Recurse -File | Measure-Object Length -Sum).Sum / 1MB
Write-Host ("App folder size: {0:N1} MB" -f $size)
}   # end of: if (-not $InstallerOnly)

# ---- 5. installer (Inno Setup), optional
if ($SkipInstaller) { Write-Host "`nSkipped installer (-SkipInstaller)."; exit 0 }
Step "Installer"
$iscc = @(
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) {
    Write-Host "Inno Setup 6 not found - installer skipped. Install it (winget install JRSoftware.InnoSetup) and re-run." -ForegroundColor Yellow
    exit 0
}
if (-not (Test-Path (Join-Path $dist "CUPL-Undergraduate-Thesis-Auto-Formatter\CUPL-Undergraduate-Thesis-Auto-Formatter.exe"))) { Fail "dist\CUPL-Undergraduate-Thesis-Auto-Formatter not found - run without -InstallerOnly first" }
# the wizard's license page shows the very same text as the in-app agreement (single source)
& $py (Join-Path $PSScriptRoot "make_license.py")
if ($LASTEXITCODE -ne 0) { Fail "make_license.py failed" }
# license notices shipped with the installer (also syncs the LICENSE file in the project root)
& $py (Join-Path $PSScriptRoot "make_notices.py")
if ($LASTEXITCODE -ne 0) { Fail "make_notices.py failed" }
$ver = & $py -c "from essay_tool import __version__; print(__version__)"
& $iscc "/DAppVersion=$ver" (Join-Path $PSScriptRoot "installer.iss")
if ($LASTEXITCODE -ne 0) { Fail "Inno Setup failed" }
$setup = Get-ChildItem (Join-Path $dist "installer") -Filter "*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
Write-Host ("Installer: {0}  ({1:N1} MB)" -f $setup.FullName, ($setup.Length / 1MB)) -ForegroundColor Green
