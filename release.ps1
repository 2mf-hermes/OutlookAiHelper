# Publishes a release: build -> tag -> GitHub Release with the artifact attached.
#
#   powershell -File release.ps1 -Repo owner/name
#   powershell -File release.ps1 -Repo owner/name -DryRun
#
# The tag, the running build's version and the asset name all come from VERSION and
# from UpdateSecurityPolicy.ExpectedAssetFileName, so they cannot drift apart.
param(
    [Parameter(Mandatory = $true)][string]$Repo,
    [string]$Notes = '',
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $root

$version = (Get-Content (Join-Path $root 'VERSION') -Raw).Trim()
if ($version -notmatch '^[0-9]+\.[0-9]+\.[0-9]+$') {
    throw "VERSION must look like 1.0.1 (found '$version')"
}
$tag = "v$version"
$artifact = Join-Path $root 'dist\OutlookAiHelper.exe'

Write-Host "Releasing $tag from $Repo"

# --- Preflight ---------------------------------------------------------------
if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { throw 'gh CLI not found' }
if (-not (Test-Path $artifact)) { throw "Missing artifact: $artifact (run build.ps1 first)" }

# The published tag must match the version the downloaded build will report.
$informed = (Get-Item $artifact).VersionInfo.ProductVersion
if ($informed -and $informed -notlike "$version*") {
    throw "Artifact reports $informed but VERSION says $version - rebuild first (build.ps1)"
}

$status = (git status --porcelain)
if ($status) { throw "Uncommitted changes:\n$status\nCommit them before releasing." }

& git rev-parse -q --verify "refs/tags/$tag" | Out-Null
if ($LASTEXITCODE -eq 0) { throw "Tag $tag already exists locally" }

if ($DryRun) {
    Write-Host "[dry run] would: git tag $tag; push; gh release create $tag --repo $Repo --title $tag --notes ... $artifact"
    exit 0
}

# --- Build, tag, push, publish ----------------------------------------------
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root 'build.ps1')
if ($LASTEXITCODE -ne 0) { throw 'build.ps1 failed' }

& git tag $tag
if ($LASTEXITCODE -ne 0) { throw "git tag $tag failed" }

& git push origin HEAD
if ($LASTEXITCODE -ne 0) { throw 'git push (branch) failed' }

& git push origin $tag
if ($LASTEXITCODE -ne 0) { throw "git push $tag failed" }

if ([string]::IsNullOrWhiteSpace($Notes)) {
    $Notes = "Outlook AI 助手 $version`n`n以內建更新檢查下載安裝，或直接取用本頁附檔。"
}

& gh release create $tag $artifact --repo $Repo --title $tag --notes $Notes
if ($LASTEXITCODE -ne 0) { throw 'gh release create failed' }

Write-Host "Released $tag -> https://github.com/$Repo/releases/tag/$tag"
