param(
  [string]$ManifestPath = "manifests/manifest.staging.xml",
  [string]$SharePath = "$env:USERPROFILE\OfficeAddinCatalog"
)

$ErrorActionPreference = 'Stop'

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$resolvedManifest = Resolve-Path (Join-Path $projectRoot $ManifestPath)

if (-not (Test-Path $SharePath)) {
  throw "Share path does not exist: $SharePath. Run setup-local-share.ps1 first."
}

$manifestFileName = Split-Path $resolvedManifest -Leaf
$destination = Join-Path $SharePath $manifestFileName
Copy-Item -Path $resolvedManifest -Destination $destination -Force

Write-Host "Published manifest to share catalog:"
Write-Host "  $destination"
Write-Host ""
if ($manifestFileName -match 'outlook') {
  Write-Host "To sideload in Outlook: File > Manage Add-ins, or via the Add-ins dialog in Outlook Desktop"
} else {
  Write-Host "To sideload in each Office app:"
  Write-Host "  Excel       : Home > Add-ins > More Add-ins > Shared Folder"
  Write-Host "  PowerPoint  : Insert > Add-ins > More Add-ins > Shared Folder"
  Write-Host "  Word        : Insert > Add-ins > More Add-ins > Shared Folder"
  Write-Host "  (For Outlook, run: npm run sideload:share:publish:outlook)"
}
