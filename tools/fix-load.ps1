<#
    PExL add-in load helper
    -----------------------
    The add-in now targets .NET Framework 4.8. On the desktop CLR, Excel-DNA loads
    the managed add-in in-process and does NOT extract a .NET host DLL into %TEMP%,
    so it no longer trips the Defender Exploit Guard / ASR rule that caused
    "Loading ExcelDna.ManagedHost failed: 0x80070005" with the old .NET 8 build.

    You usually do NOT need admin for the net48 build. This script just:
      1. Unblocks every file in the repo (removes Mark-of-the-Web).
      2. (Optional, admin) Adds Defender exclusions as belt-and-suspenders.
      3. Prints the exact file you should load in Excel for your bitness.

    Run normally, or elevated if you want step 2:

        powershell -ExecutionPolicy Bypass -File "C:\dev\PExL\tools\fix-load.ps1"
#>

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
Write-Host "PExL repo: $repo" -ForegroundColor Cyan

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
}

# 1. Unblock everything (safe, no admin needed) ------------------------------
Write-Host "`n[1/3] Unblocking files (removing Mark-of-the-Web)..." -ForegroundColor Yellow
Get-ChildItem $repo -Recurse -File | Unblock-File
Write-Host "      done." -ForegroundColor Green

# 2. Defender exclusions (needs admin) ---------------------------------------
Write-Host "`n[2/3] Adding Windows Defender exclusions..." -ForegroundColor Yellow
$excelDnaTemp = Join-Path $env:LOCALAPPDATA "Temp\Excel-DNA"
if (Test-Admin) {
    try {
        Add-MpPreference -ExclusionPath $repo -ErrorAction Stop
        Add-MpPreference -ExclusionPath $excelDnaTemp -ErrorAction SilentlyContinue
        Add-MpPreference -ExclusionProcess "EXCEL.EXE" -ErrorAction SilentlyContinue
        Write-Host "      Excluded: $repo" -ForegroundColor Green
        Write-Host "      Excluded: $excelDnaTemp" -ForegroundColor Green
    } catch {
        Write-Host "      Could not add Defender exclusion: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "      Your machine may use a managed (corporate) Defender policy." -ForegroundColor Red
    }
} else {
    Write-Host "      SKIPPED - not running as administrator." -ForegroundColor Red
    Write-Host "      Re-run this script as administrator to add the exclusion." -ForegroundColor Red
}

# 3. Tell the user which xll to load -----------------------------------------
Write-Host "`n[3/3] Pick the correct add-in for your Excel bitness:" -ForegroundColor Yellow
$platform = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue).Platform

# The output folder differs depending on how it was built. Prefer the net48
# (.NET Framework) build; fall back to the legacy net8 output if that's all
# that exists. Pick whichever actually contains a freshly built xll.
$binCandidates = @(
    (Join-Path $repo "src\PExL.AddIn\bin\x64\Release\net48"),
    (Join-Path $repo "src\PExL.AddIn\bin\Release\net48"),
    (Join-Path $repo "src\PExL.AddIn\bin\x64\Release\net8.0-windows"),
    (Join-Path $repo "src\PExL.AddIn\bin\Release\net8.0-windows")
) | Where-Object { Test-Path (Join-Path $_ "PExL64.xll") }
$bin = $binCandidates | Sort-Object { (Get-Item (Join-Path $_ "PExL64.xll")).LastWriteTime } -Descending | Select-Object -First 1
if (-not $bin) {
    Write-Host "      No built add-in found. Run:  dotnet build src\PExL.AddIn\PExL.AddIn.csproj -c Release -p:Platform=x64" -ForegroundColor Red
    return
}
if ($platform -eq 'x64' -or [Environment]::Is64BitOperatingSystem) {
    Write-Host "      Excel appears to be 64-bit." -ForegroundColor Green
    Write-Host "      Load this file in Excel (File > Options > Add-ins > Manage: Excel Add-ins > Browse):" -ForegroundColor Cyan
    Write-Host "        $bin\PExL64.xll" -ForegroundColor White
} else {
    Write-Host "      Excel appears to be 32-bit." -ForegroundColor Green
    Write-Host "        $bin\PExL.xll" -ForegroundColor White
}

Write-Host "`n      NOTE: the .xll needs its sibling files to run the editor:" -ForegroundColor DarkGray
Write-Host "            PExL.Core.dll, Newtonsoft.Json.dll, Microsoft.Web.WebView2.*.dll," -ForegroundColor DarkGray
Write-Host "            WebView2Loader.dll, and the web\ folder." -ForegroundColor DarkGray
Write-Host "            To share with testers, zip the whole '$([IO.Path]::GetFileName($bin))' folder." -ForegroundColor DarkGray

Write-Host "`nDone. Fully close ALL Excel windows, then load the file above." -ForegroundColor Cyan
