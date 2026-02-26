param(
  [string]$ManifestPath = "manifests/manifest.dev.xml",
  [switch]$All,
  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

function Resolve-FullPath([string]$PathValue) {
  if ([System.IO.Path]::IsPathRooted($PathValue)) {
    return [System.IO.Path]::GetFullPath($PathValue)
  }
  $repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
  return [System.IO.Path]::GetFullPath((Join-Path $repoRoot $PathValue))
}

$regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
if (-not (Test-Path $regPath)) {
  Write-Host "No Office developer sideload entries found."
  exit 0
}

if ($All) {
  if ($DryRun) {
    Write-Host "[DryRun] Would remove all entries under $regPath"
  } else {
    Remove-Item -Path $regPath -Recurse -Force
    Write-Host "Removed all sideloaded add-in entries."
  }
  exit 0
}

$manifestFullPath = Resolve-FullPath $ManifestPath
$props = (Get-ItemProperty -Path $regPath)
$removed = 0

foreach ($p in $props.PSObject.Properties) {
  if ($p.Name -notmatch '^\d+$') {
    continue
  }
  if ([string]$p.Value -eq $manifestFullPath) {
    if ($DryRun) {
      Write-Host "[DryRun] Would remove index $($p.Name) => $manifestFullPath"
    } else {
      Remove-ItemProperty -Path $regPath -Name $p.Name -Force
      Write-Host "Removed index $($p.Name)."
    }
    $removed++
  }
}

if ($removed -eq 0) {
  Write-Host "No matching sideload entry found for: $manifestFullPath"
} else {
  Write-Host "Done. Removed $removed entr$(if($removed -eq 1){'y'}else{'ies'})."
}
