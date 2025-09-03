# Build Installer Script for BulkEditor
# This script builds the application and creates an installer package

param(
    [Parameter(Mandatory=$false)]
    [string]$Configuration = "Release",

    [Parameter(Mandatory=$false)]
    [string]$Version = "1.0.0.0",

    [Parameter(Mandatory=$false)]
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot
$outputDir = Join-Path $rootDir "Output"
$installerDir = Join-Path $rootDir "BulkEditor.Installer"

Write-Host "Building BulkEditor Installer" -ForegroundColor Green
Write-Host "Configuration: $Configuration" -ForegroundColor Yellow
Write-Host "Version: $Version" -ForegroundColor Yellow
Write-Host "Root Directory: $rootDir" -ForegroundColor Yellow

# Clean output directory
if (Test-Path $outputDir) {
    Write-Host "Cleaning output directory..." -ForegroundColor Yellow
    Remove-Item $outputDir -Recurse -Force
}
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

# Build the application if not skipping
if (-not $SkipBuild) {
    Write-Host "Building application..." -ForegroundColor Yellow

    # Restore packages
    Write-Host "Restoring NuGet packages..." -ForegroundColor Cyan
    dotnet restore "$rootDir\BulkEditor.sln"
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to restore packages"
    }

    # Build solution
    Write-Host "Building solution in $Configuration mode..." -ForegroundColor Cyan
    dotnet build "$rootDir\BulkEditor.sln" --configuration $Configuration --no-restore
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to build solution"
    }

    # Publish the UI project
    Write-Host "Publishing UI project..." -ForegroundColor Cyan
    $publishDir = Join-Path $outputDir "Publish"
    dotnet publish "$rootDir\BulkEditor.UI\BulkEditor.UI.csproj" `
        --configuration $Configuration `
        --output $publishDir `
        --no-restore
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to publish UI project"
    }
}

# Check if WiX is available
Write-Host "Checking for WiX Toolset..." -ForegroundColor Yellow
$wixPath = Get-Command "dotnet" -ErrorAction SilentlyContinue
if (-not $wixPath) {
    throw "dotnet CLI not found. Please install .NET SDK."
}

# Build installer using WiX
Write-Host "Building installer..." -ForegroundColor Yellow
try {
    Push-Location $installerDir

    # Set version in project file if specified
    if ($Version -ne "1.0.0.0") {
        Write-Host "Updating version to $Version..." -ForegroundColor Cyan
        $projectContent = Get-Content "BulkEditor.Installer.wixproj" -Raw
        $projectContent = $projectContent -replace "<ProductVersion>.*</ProductVersion>", "<ProductVersion>$Version</ProductVersion>"
        Set-Content "BulkEditor.Installer.wixproj" -Value $projectContent
    }

    # Build the installer
    dotnet build --configuration $Configuration
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to build installer"
    }

    # Copy installer to output directory
    $msiFile = Get-ChildItem "bin\$Configuration" -Filter "*.msi" | Select-Object -First 1
    if ($msiFile) {
        $outputMsi = Join-Path $outputDir "BulkEditor-Setup-$Version.msi"
        Copy-Item $msiFile.FullName $outputMsi
        Write-Host "Installer created: $outputMsi" -ForegroundColor Green
    } else {
        throw "Installer MSI file not found"
    }
}
finally {
    Pop-Location
}

# Create version info file
Write-Host "Creating version information..." -ForegroundColor Yellow
$versionInfo = @{
    Version = $Version
    BuildDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Configuration = $Configuration
    Installer = "BulkEditor-Setup-$Version.msi"
} | ConvertTo-Json -Depth 2

Set-Content (Join-Path $outputDir "version.json") -Value $versionInfo

Write-Host "Build completed successfully!" -ForegroundColor Green
Write-Host "Output directory: $outputDir" -ForegroundColor Green

# Display summary
Write-Host "`nBuild Summary:" -ForegroundColor Cyan
Write-Host "- Application built and published" -ForegroundColor White
Write-Host "- Installer created: BulkEditor-Setup-$Version.msi" -ForegroundColor White
Write-Host "- Version information saved" -ForegroundColor White
Write-Host "`nReady for distribution!" -ForegroundColor Green