# Set variables
$repo = "ItMeDiaTech/Bulk_Editor_WPF"
$releasesUrl = "https://api.github.com/repos/$repo/releases/latest"
$releasesFolder = "$PSScriptRoot\Releases"

# Get latest release info
$release = Invoke-RestMethod -Uri $releasesUrl

# Find the first asset with .msi extension (update if you want a different type)
$asset = $release.assets | Where-Object { $_.name -like "*.msi" } | Select-Object -First 1

if ($null -eq $asset) {
    Write-Host "No MSI asset found in the latest release."
    exit 1
}

# Ensure Releases folder exists
if (-not (Test-Path $releasesFolder)) {
    New-Item -ItemType Directory -Path $releasesFolder | Out-Null
}

# Download the asset
$downloadPath = Join-Path $releasesFolder $asset.name
Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $downloadPath

Write-Host "Downloaded $($asset.name) to $releasesFolder"