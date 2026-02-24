param(
  [string]$ManifestPath = "manifests/manifest.dev.xml",
  [switch]$SkipCertTrust,
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

function Add-ManifestRegistration([string]$ManifestFullPath) {
  $regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

  if (-not (Test-Path $regPath)) {
    if ($DryRun) {
      Write-Host "[DryRun] Would create $regPath"
    } else {
      New-Item -Path $regPath -Force | Out-Null
    }
  }

  $existing = @{}
  if (Test-Path $regPath) {
    $props = (Get-ItemProperty -Path $regPath)
    foreach ($p in $props.PSObject.Properties) {
      if ($p.Name -match '^\d+$') {
        $existing[$p.Name] = [string]$p.Value
      }
    }
  }

  foreach ($entry in $existing.GetEnumerator()) {
    if ($entry.Value -eq $ManifestFullPath) {
      Write-Host "Manifest already registered at index $($entry.Key)."
      return
    }
  }

  $nextIndex = 0
  while ($existing.ContainsKey([string]$nextIndex)) {
    $nextIndex++
  }

  if ($DryRun) {
    Write-Host "[DryRun] Would register manifest at index $nextIndex => $ManifestFullPath"
  } else {
    New-ItemProperty -Path $regPath -Name ([string]$nextIndex) -Value $ManifestFullPath -PropertyType String -Force | Out-Null
    Write-Host "Manifest registered at index $nextIndex."
  }
}

function Trust-CertificateIfAvailable {
  if ($SkipCertTrust) {
    Write-Host "Skipping certificate trust (SkipCertTrust)."
    return
  }

  $certCandidates = @(
    (Join-Path $env:USERPROFILE '.office-addin-dev-certs\localhost.crt'),
    (Join-Path $env:USERPROFILE '.office-addin-dev-certs\localhost.pem')
  )

  $certPath = $certCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
  if (-not $certPath) {
    Write-Host "No local dev certificate found at ~/.office-addin-dev-certs; skipping trust step."
    return
  }

  Write-Host "Using certificate: $certPath"
  if ($DryRun) {
    Write-Host "[DryRun] Would import certificate to Cert:\CurrentUser\Root"
    return
  }

  try {
    Import-Certificate -FilePath $certPath -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
    Write-Host "Certificate trusted in CurrentUser Root store."
  } catch {
    Write-Host "Certificate trust failed: $($_.Exception.Message)"
    Write-Host "Continue anyway; Office may still prompt on first load."
  }
}

$manifestFullPath = Resolve-FullPath $ManifestPath
if (-not (Test-Path $manifestFullPath)) {
  throw "Manifest not found: $manifestFullPath"
}

Write-Host "Registering Office add-in (Word, PowerPoint, Excel only)..."
Write-Host "Manifest: $manifestFullPath"

Trust-CertificateIfAvailable
Add-ManifestRegistration -ManifestFullPath $manifestFullPath

Write-Host "Done."
Write-Host "Next: open Word, PowerPoint, or Excel and add via Insert > Add-ins > My Add-ins."
