<#
.SYNOPSIS
    Builds the packed PExL add-in locally and assembles a ready-to-ship .zip.

.DESCRIPTION
    Mirrors what .github/workflows/release.yml does, but on your machine — handy
    for smoke-testing the exact bundle before you tag a release.

    Produces:  dist/PExL-<version>/      (the unzipped staging folder)
               dist/PExL-<version>.zip    (full package — corporate/restricted)
               dist/PExL-<version>.xll    (lightweight single file — broader audience)

.PARAMETER Version
    Version string, e.g. 0.1.0 or v0.1.0. Defaults to 0.0.0-dev.

.PARAMETER Tag
    If supplied, after a successful build the script creates an annotated git
    tag (v<version>) and pushes it to origin — which triggers the Release
    workflow on GitHub to build and publish the official artifact.

.EXAMPLE
    pwsh tools/make-release.ps1 -Version 0.1.0

.EXAMPLE
    pwsh tools/make-release.ps1 -Version 0.1.0 -Tag
#>
[CmdletBinding()]
param(
    [string]$Version = '0.0.0-dev',
    [switch]$Tag
)

$ErrorActionPreference = 'Stop'

# Run from the repo root regardless of where the script is invoked.
$repo = Split-Path -Parent $PSScriptRoot
Set-Location $repo

$ver = $Version.TrimStart('v')
$tagName = "v$ver"
Write-Host "==> Building PExL $tagName" -ForegroundColor Cyan

dotnet restore PExL.sln
dotnet build PExL.sln -c Release -p:Platform=x64 --no-restore
dotnet test tests/PExL.Core.Tests/PExL.Core.Tests.csproj -c Release --no-build --verbosity quiet

$out = 'src/PExL.AddIn/bin/x64/Release/net48'
$xll = Get-ChildItem -Path $out -Recurse -Filter '*64-packed.xll' | Select-Object -First 1
if (-not $xll) { throw "Packed 64-bit .xll not found under '$out'. Did the build/pack run?" }

$name  = "PExL-$tagName"
$stage = Join-Path 'dist' $name
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force -Path $stage | Out-Null

Copy-Item $xll.FullName (Join-Path $stage 'PExL.xll')

$wv2 = Join-Path $out 'WebView2Loader.dll'
if (Test-Path $wv2) { Copy-Item $wv2 $stage } else { Write-Warning 'WebView2Loader.dll not found' }

$web = Join-Path $out 'web'
if (Test-Path $web) { Copy-Item $web (Join-Path $stage 'web') -Recurse } else { Write-Warning 'web\ folder not found' }

if (Test-Path 'README.html')        { Copy-Item 'README.html' $stage }
if (Test-Path 'release/INSTALL.txt') { Copy-Item 'release/INSTALL.txt' $stage }

# Artifact 1: full ZIP (corporate/restricted — everything works).
$zip = Join-Path 'dist' "$name.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip -Force

# Artifact 2: standalone lightweight .xll (broader audience — ribbon + transpiler;
# the editor/docs panes need the ZIP).
$singleXll = Join-Path 'dist' "$name.xll"
Copy-Item $xll.FullName $singleXll -Force

Write-Host ""
Write-Host "==> Done. Artifacts:" -ForegroundColor Green
Write-Host "    $((Resolve-Path $zip).Path)        (full package)"
Write-Host "    $((Resolve-Path $singleXll).Path)  (lightweight single file)"
Write-Host "    ZIP contents:"
Get-ChildItem -Recurse $stage | ForEach-Object { "      $($_.FullName.Substring((Resolve-Path $stage).Path.Length + 1))" }

if ($Tag) {
    Write-Host ""
    Write-Host "==> Tagging $tagName and pushing to origin (triggers the Release workflow)" -ForegroundColor Cyan
    git tag -a $tagName -m "PExL $tagName"
    git push origin $tagName
    Write-Host "    Pushed. Watch the build at: https://github.com/keeylogger/PExL/actions"
}
