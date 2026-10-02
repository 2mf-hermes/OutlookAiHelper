$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $root 'src\OutlookAiHelper'
$tests = Join-Path $root 'tests'
$out = Join-Path $root 'dist'
New-Item -ItemType Directory -Force -Path $out | Out-Null

# --- Version stamping --------------------------------------------------------
# VERSION is the single source of truth. It feeds the assembly identity the app
# reports at runtime, the git tag of the release the updater compares against,
# and the name of the artifact uploaded to that release. This script only reads.
$versionFile = Join-Path $root 'VERSION'
if (-not (Test-Path $versionFile)) { throw "Missing VERSION file: $versionFile" }
$version = (Get-Content $versionFile -Raw).Trim()
if ($version -notmatch '^[0-9]+\.[0-9]+\.[0-9]+$') {
  throw "VERSION must look like 1.0.1 (found '$version')"
}

$objDir = Join-Path $src 'obj'
New-Item -ItemType Directory -Force -Path $objDir | Out-Null
$assemblyInfo = Join-Path $objDir 'AssemblyInfo.g.cs'
Set-Content -Path $assemblyInfo -Encoding UTF8 -Value @(
  'using System.Reflection;'
  "[assembly: AssemblyVersion(`"$version.0`")]"
  "[assembly: AssemblyFileVersion(`"$version.0`")]"
  "[assembly: AssemblyInformationalVersion(`"$version`")]"
)
Write-Host "Version $version (stamped into $assemblyInfo)"

$fw = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
if (-not (Test-Path (Join-Path $fw 'csc.exe'))) {
  $fw = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319'
}
$csc = Join-Path $fw 'csc.exe'
$wpfDir = Join-Path $fw 'WPF'

$wpf = @(
  (Join-Path $wpfDir 'PresentationFramework.dll'),
  (Join-Path $wpfDir 'PresentationCore.dll'),
  (Join-Path $wpfDir 'WindowsBase.dll'),
  (Join-Path $fw 'System.Xaml.dll')
)
$sys = @(
  (Join-Path $fw 'System.dll'),
  (Join-Path $fw 'System.Core.dll'),
  (Join-Path $fw 'System.Runtime.Serialization.dll')
)
foreach ($p in ($wpf + $sys + @($csc))) {
  if (-not (Test-Path $p)) { throw "Missing build reference: $p" }
}

$appSources = Get-ChildItem -Path $src -Filter *.cs -Recurse | Select-Object -ExpandProperty FullName
$refs = @()
foreach ($r in ($wpf + $sys)) { $refs += "/r:`"$r`"" }

$appExe = Join-Path $out 'OutlookAiHelper.exe'
$iconPath = Join-Path $src 'app.ico'
$iconArg = @()
if (Test-Path $iconPath) {
  $iconArg = "/win32icon:`"$iconPath`""
  Copy-Item $iconPath (Join-Path $out 'app.ico') -Force
}
Write-Host "Building app -> $appExe"
& $csc /nologo /target:winexe /platform:anycpu /optimize+ `
  /out:"$appExe" `
  $iconArg `
  $refs `
  $appSources
if ($LASTEXITCODE -ne 0) { throw "App build failed" }

$testExe = Join-Path $out 'OutlookAiHelper.Tests.exe'
$testSources = @($appSources) + (Get-ChildItem -Path $tests -Filter *.cs -Recurse | Select-Object -ExpandProperty FullName)
# Exclude Program.cs entry from test host to avoid duplicate Main
$testSources = $testSources | Where-Object { $_ -notmatch '\\Program\.cs$' }

Write-Host "Building tests -> $testExe"
& $csc /nologo /target:exe /platform:anycpu /optimize+ `
  /out:"$testExe" `
  $refs `
  $testSources
if ($LASTEXITCODE -ne 0) { throw "Test build failed" }

Write-Host "Running tests"
& $testExe
if ($LASTEXITCODE -ne 0) { throw "Tests failed" }

Write-Host "Build OK: $appExe"
