# Build a portable Windows bundle with a private Python runtime.
# The bundle includes MarkItDown Handy, Python packages, and launchers.

param(
    [string]$PythonVersion = $env:PY_VERSION,
    [string]$AppVersion = "0.1.0"
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($PythonVersion)) {
    $PythonVersion = "3.11"
}

$RootDir = Resolve-Path (Join-Path $PSScriptRoot "..")
$Src = Join-Path $RootDir "src\markitdown_handy.py"
$BuildDir = Join-Path $RootDir "build-windows"
$OutDir = Join-Path $RootDir "dist-windows"
$BundleName = "MarkItDown Handy Portable"
$BundleDir = Join-Path $OutDir $BundleName
$RuntimeDir = Join-Path $BundleDir "runtime"
$AppDir = Join-Path $BundleDir "app"
$ToolsDir = Join-Path $BundleDir "tools"
$TessdataDir = Join-Path $BundleDir "tessdata"
$Arch = if ([Environment]::Is64BitOperatingSystem) { "x64" } else { "x86" }

Remove-Item -Recurse -Force $BundleDir -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $BuildDir, $OutDir, $RuntimeDir, $AppDir, $ToolsDir, $TessdataDir | Out-Null
Copy-Item $Src (Join-Path $AppDir "markitdown_handy.py") -Force

$VenvDir = Join-Path $BuildDir "venv"
if (-not (Test-Path $VenvDir)) {
    py -$PythonVersion -m venv $VenvDir
}

$Python = Join-Path $VenvDir "Scripts\python.exe"
if (-not (Test-Path $Python)) {
    $Python = "python"
}

& $Python -m pip install --upgrade pip
& $Python -m pip install --upgrade "markitdown[all]" tkinterdnd2 ocrmypdf

& $Python - <<'PYCHECK'
import importlib.util
missing = [m for m in ["markitdown", "tkinterdnd2", "ocrmypdf"] if not importlib.util.find_spec(m)]
if missing:
    raise SystemExit(f"Missing required modules: {missing}")
print("Python module check OK")
PYCHECK

Copy-Item -Recurse -Force (Join-Path $VenvDir "*") $RuntimeDir

$RuntimePython = Join-Path $RuntimeDir "Scripts\python.exe"
$RuntimeScripts = Join-Path $RuntimeDir "Scripts"

$LauncherCmd = Join-Path $BundleDir "MarkItDown Handy.cmd"
@"
@echo off
setlocal
set "APP_DIR=%~dp0"
set "MARKITDOWN_BUNDLED_ENV=%APP_DIR%runtime"
set "MARKITDOWN_RESOURCE_DIR=%APP_DIR%"
set "PYTHONNOUSERSITE=1"
set "PATH=%APP_DIR%tools;%APP_DIR%runtime\Scripts;%APP_DIR%runtime;%PATH%"
if exist "%APP_DIR%tessdata" set "TESSDATA_PREFIX=%APP_DIR%tessdata"
"%APP_DIR%runtime\Scripts\python.exe" "%APP_DIR%app\markitdown_handy.py"
endlocal
"@ | Set-Content -Encoding ASCII $LauncherCmd

$LauncherPs1 = Join-Path $BundleDir "MarkItDown Handy.ps1"
@"
`$AppDir = Split-Path -Parent `$MyInvocation.MyCommand.Path
`$env:MARKITDOWN_BUNDLED_ENV = Join-Path `$AppDir 'runtime'
`$env:MARKITDOWN_RESOURCE_DIR = `$AppDir
`$env:PYTHONNOUSERSITE = '1'
`$env:PATH = (Join-Path `$AppDir 'tools') + ';' + (Join-Path `$AppDir 'runtime\Scripts') + ';' + (Join-Path `$AppDir 'runtime') + ';' + `$env:PATH
`$Tessdata = Join-Path `$AppDir 'tessdata'
if (Test-Path `$Tessdata) { `$env:TESSDATA_PREFIX = `$Tessdata }
& (Join-Path `$AppDir 'runtime\Scripts\python.exe') (Join-Path `$AppDir 'app\markitdown_handy.py')
"@ | Set-Content -Encoding UTF8 $LauncherPs1

# Put Windows-friendly command shims in runtime\bin because the app's bundled-runtime lookup
# was originally written for macOS conda-pack bundles.
$RuntimeBin = Join-Path $RuntimeDir "bin"
New-Item -ItemType Directory -Force $RuntimeBin | Out-Null
@"
@echo off
"%~dp0..\Scripts\python.exe" -m markitdown %*
"@ | Set-Content -Encoding ASCII (Join-Path $RuntimeBin "markitdown.bat")
@"
@echo off
"%~dp0..\Scripts\python.exe" -m ocrmypdf %*
"@ | Set-Content -Encoding ASCII (Join-Path $RuntimeBin "ocrmypdf.bat")
Copy-Item -Force (Join-Path $RuntimeScripts "python.exe") (Join-Path $RuntimeBin "python.exe")

# Best-effort copy of external OCR/PDF tools available on the build runner.
# The app still works for direct MarkItDown conversion without these tools.
$ToolNames = @(
    "tesseract.exe",
    "gswin64c.exe",
    "gswin32c.exe",
    "qpdf.exe"
)
foreach ($tool in $ToolNames) {
    $cmd = Get-Command $tool -ErrorAction SilentlyContinue
    if ($cmd -and (Test-Path $cmd.Source)) {
        Copy-Item -Force $cmd.Source $ToolsDir
        Write-Host "Copied tool: $($cmd.Source)"
    } else {
        Write-Host "Tool not found, skipping: $tool"
    }
}

$TessCandidates = @(
    "$env:ProgramFiles\Tesseract-OCR\tessdata",
    "$env:ProgramFiles(x86)\Tesseract-OCR\tessdata",
    "$env:ChocolateyInstall\lib\tesseract\tools\tessdata"
)
foreach ($candidate in $TessCandidates) {
    if (Test-Path $candidate) {
        foreach ($lang in @("eng", "osd", "chi_sim", "chi_tra")) {
            $srcLang = Join-Path $candidate "$lang.traineddata"
            if (Test-Path $srcLang) {
                Copy-Item -Force $srcLang $TessdataDir
            }
        }
    }
}

& $RuntimePython - <<'PYCHECK'
import importlib.util
for mod in ["markitdown", "tkinterdnd2", "ocrmypdf"]:
    if not importlib.util.find_spec(mod):
        raise SystemExit(f"Bundled module missing: {mod}")
print("Bundled runtime check OK")
PYCHECK

$ZipPath = Join-Path $OutDir "MarkItDown_Handy_v${AppVersion}_portable_Windows_${Arch}.zip"
Remove-Item -Force $ZipPath -ErrorAction SilentlyContinue
Compress-Archive -Path $BundleDir -DestinationPath $ZipPath -Force

Write-Host "Built Windows portable bundle: $BundleDir"
Write-Host "Zip: $ZipPath"
