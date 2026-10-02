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
$assetName = 'OutlookAiHelper.exe'   # UpdateSecurityPolicy.ExpectedAssetFileName
$artifact = Join-Path $root "dist\$assetName"

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
if ($status) { throw "Uncommitted changes:`n$status`nCommit them before releasing." }

$head = (git rev-parse HEAD).Trim()
$localTag = (& git rev-parse -q --verify "refs/tags/$tag")
if ($localTag -and $localTag.Trim() -ne $head) {
    throw "Tag $tag already exists and does not point at HEAD - bump VERSION instead"
}

# gh reads '#' in an argument as the file/display-name separator, so a work folder
# containing one (as this machine's session path does) would be split in half.
# Upload a copy from a path that cannot contain one.
$uploadDir = Join-Path $env:TEMP 'OutlookAiHelper-release'
New-Item -ItemType Directory -Force -Path $uploadDir | Out-Null
$upload = Join-Path $uploadDir $assetName

if ($DryRun) {
    Write-Host "[dry run] would: build; tag/push $tag; publish $upload ($((Get-Item $artifact).Length) bytes) to $Repo"
    exit 0
}

# --- Build --------------------------------------------------------------------
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root 'build.ps1')
if ($LASTEXITCODE -ne 0) { throw 'build.ps1 failed' }
Copy-Item $artifact $upload -Force

# --- Tag ----------------------------------------------------------------------
if (-not $localTag) {
    & git tag $tag
    if ($LASTEXITCODE -ne 0) { throw "git tag $tag failed" }
}

& git push origin HEAD
if ($LASTEXITCODE -ne 0) { throw 'git push (branch) failed' }

$remoteTag = (& git ls-remote --tags origin "refs/tags/$tag")
if (-not $remoteTag) {
    & git push origin $tag
    if ($LASTEXITCODE -ne 0) { throw "git push $tag failed" }
}

# --- Publish ------------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($Notes)) {
    $Notes = "Outlook AI 助手 $version`n`n以內建更新檢查下載安裝，或直接取用本頁附檔。"
}

& gh release view $tag --repo $Repo 1>$null 2>$null
$releaseExists = ($LASTEXITCODE -eq 0)

if ($releaseExists) {
    Write-Host "Release $tag exists - replacing the artifact"
    & gh release upload $tag $upload --repo $Repo --clobber
    if ($LASTEXITCODE -ne 0) { throw 'gh release upload failed' }
} else {
    & gh release create $tag $upload --repo $Repo --title $tag --notes $Notes
    if ($LASTEXITCODE -ne 0) { throw 'gh release create failed' }
}

# --- Verify what the updater will see ----------------------------------------
$json = (& gh release view $tag --repo $Repo --json isDraft,isPrerelease,assets)
if ($LASTEXITCODE -ne 0) { throw 'could not read the release back' }
$release = $json | ConvertFrom-Json
if ($release.isDraft -or $release.isPrerelease) {
    throw "Release $tag is $(if ($release.isDraft) { 'a draft' } else { 'a prerelease' }) - the updater ignores both"
}
$names = @($release.assets | ForEach-Object { $_.name })
if ($names -notcontains $assetName) {
    throw "Release $tag has no asset named $assetName (found: $($names -join ', '))"
}

Write-Host "Released $tag with $assetName -> https://github.com/$Repo/releases/tag/$tag"
