# Detail-panel harness: builds it, opens the real window in-process, drives the panel's own entry
# points, and reports what the panel rendered.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File selftest\run_panel.ps1
#
# No arguments. It runs against dist\OutlookAiHelper.exe, so run build.ps1 first.
#
# Why in-process: the right-hand panel is a collapsed column until a row is clicked, and a synthetic
# click needs a window in the foreground, which an unattended machine will not give (the UI
# Automation smoke test skips its click checks for the same reason). Calling the window's handlers
# on the UI thread is the one route that renders the panel here. See PanelHarness.cs for the four
# passes and selftest\panel\panel.txt for what each one asserts.
$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
$data = Join-Path $here 'panel\panel.txt'
$outTxt = Join-Path $here 'panel-out.txt'
$outPng = Join-Path $here 'panel-out.png'

$appExe = Join-Path $root 'dist\OutlookAiHelper.exe'
if (-not (Test-Path -LiteralPath $appExe)) { throw "Missing app build: $appExe - run build.ps1 first." }
$stale = Get-ChildItem -LiteralPath (Join-Path $root 'src\OutlookAiHelper') -Filter '*.cs' -Recurse |
    Where-Object { $_.LastWriteTime -gt (Get-Item -LiteralPath $appExe).LastWriteTime }
if ($stale) {
    Write-Warning ("dist\OutlookAiHelper.exe is older than " + $stale.Count +
        " source file(s): run build.ps1 or this run grades a stale binary.")
}

$fw = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
if (-not (Test-Path -LiteralPath (Join-Path $fw 'csc.exe'))) {
    $fw = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319'
}
$csc = Join-Path $fw 'csc.exe'
$wpfDir = Join-Path $fw 'WPF'

# Old output must not survive a failed run: a leftover PNG is exactly how a screenshot lies.
Remove-Item -LiteralPath $outTxt -Force -ErrorAction SilentlyContinue
Get-ChildItem -LiteralPath $here -Filter 'panel-out*.png' -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue

# Build in the temp folder, not in the repo. The harness exe is written beside a copy of the app, so
# the window type it reflects on is the same assembly its TodoItem comes from.
$stage = Join-Path $env:TEMP 'oah-panel-harness'
New-Item -ItemType Directory -Force -Path $stage | Out-Null
Copy-Item -LiteralPath $appExe -Destination (Join-Path $stage 'OutlookAiHelper.exe') -Force

# Bare /r: tokens: PowerShell quotes each whole argument itself, and hand-quoting splits any path
# that contains a space.
$refs = @(
    ('/r:' + (Join-Path $wpfDir 'PresentationFramework.dll')),
    ('/r:' + (Join-Path $wpfDir 'PresentationCore.dll')),
    ('/r:' + (Join-Path $wpfDir 'WindowsBase.dll')),
    ('/r:' + (Join-Path $fw 'System.Xaml.dll')),
    ('/r:' + (Join-Path $fw 'System.dll')),
    ('/r:' + (Join-Path $fw 'System.Core.dll')),
    ('/r:' + (Join-Path $stage 'OutlookAiHelper.exe'))
)

$source = Join-Path $here 'PanelHarness.cs'
$exe = Join-Path $stage 'PanelHarness.exe'
Write-Host "compiling harness -> $exe"
# /codepage:65001 because these sources carry UTF-8 without a BOM.
& $csc /nologo /codepage:65001 /target:exe /platform:anycpu $refs ('/out:' + $exe) $source
if ($LASTEXITCODE -ne 0) { throw 'Harness build failed' }

Write-Host 'running harness'
& $exe $outTxt $outPng $data
$code = $LASTEXITCODE

# The panel's text is CJK, so the console cannot be trusted to show it: read the log back as UTF-8
# and report counts, leaving the wording to the file.
$lines = Get-Content -LiteralPath $outTxt -Encoding UTF8
$ok = 0; $total = 0; $hard = 0
foreach ($l in $lines) {
    if ($l -match '^  [A-Za-z0-9_.]+: (\d+)/(\d+)$') { $ok += [int]$Matches[1]; $total += [int]$Matches[2] }
    if ($l -match 'FATAL|HARNESS ERROR') { $hard++ }
}

Write-Host ''
Write-Host "panel: $ok/$total assertions on screen, $hard hard failure(s), harness exit $code"
Write-Host "what the panel rendered: $outTxt"
Write-Host "screenshots:            $here\panel-out-*.png"
if ($hard -gt 0) { Write-Host 'a pass could not be driven to the end: read the log for FATAL lines.' }
exit $code
