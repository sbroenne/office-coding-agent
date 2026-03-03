param(
  [string]$CatalogUrl = "\\$env:COMPUTERNAME\OfficeAddinCatalog"
)

$ErrorActionPreference = 'Stop'

$catalogsRoot = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs'
$appRoot = 'HKCU:\Software\office-coding-agent'

if (-not (Test-Path $catalogsRoot)) {
  New-Item -Path $catalogsRoot -Force | Out-Null
}

if (-not (Test-Path $appRoot)) {
  New-Item -Path $appRoot -Force | Out-Null
}

$catalogId = $null

Get-ChildItem -Path $catalogsRoot -ErrorAction SilentlyContinue | ForEach-Object {
  try {
    $url = (Get-ItemProperty -Path $_.PSPath -Name Url -ErrorAction SilentlyContinue).Url
    if ($url -eq $CatalogUrl) {
      $catalogId = $_.PSChildName
    }
  } catch {
  }
}

if (-not $catalogId) {
  $catalogId = '{' + [guid]::NewGuid().ToString() + '}'
}

$catalogKey = Join-Path $catalogsRoot $catalogId
if (-not (Test-Path $catalogKey)) {
  New-Item -Path $catalogKey -Force | Out-Null
}

Set-ItemProperty -Path $catalogKey -Name Id -Value $catalogId
Set-ItemProperty -Path $catalogKey -Name Url -Value $CatalogUrl
Set-ItemProperty -Path $catalogKey -Name Flags -Type DWord -Value 1

Set-ItemProperty -Path $appRoot -Name LocalShareCatalogId -Value $catalogId
Set-ItemProperty -Path $appRoot -Name LocalShareCatalogUrl -Value $CatalogUrl

Write-Host "Trusted catalog configured."
Write-Host "Catalog URL: $CatalogUrl"
Write-Host "Catalog ID:  $catalogId"
Write-Host ""
Write-Host "Restart any open Office apps (Excel, PowerPoint, Word), then load the add-in:"
Write-Host "  Excel       : Home > Add-ins > More Add-ins > Shared Folder"
Write-Host "  PowerPoint  : Insert > Add-ins > More Add-ins > Shared Folder"
Write-Host "  Word        : Insert > Add-ins > More Add-ins > Shared Folder"
