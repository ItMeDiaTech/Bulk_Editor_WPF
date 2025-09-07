# Deploy Release Script for BulkEditor
# This script builds, tests, and deploys a new release to GitHub

param(
    [Parameter(Mandatory=$true)]
    [string]$Version,

    [Parameter(Mandatory=$false)]
    [string]$ReleaseNotes = "",

    [Parameter(Mandatory=$false)]
    [switch]$Prerelease,

    [Parameter(Mandatory=$false)]
    [string]$GitHubToken = $env:GITHUB_TOKEN
)

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot
$outputDir = Join-Path $rootDir "Output"

Write-Host "Deploying BulkEditor Release $Version" -ForegroundColor Green

# Load GitHub token from secure file if not provided
if ([string]::IsNullOrEmpty($GitHubToken)) {
    Write-Host "Loading GitHub token from secure file..." -ForegroundColor Yellow
    & "$PSScriptRoot\set-github-token.ps1"
    $GitHubToken = $env:GITHUB_TOKEN
}

# Validate inputs
if ([string]::IsNullOrEmpty($GitHubToken)) {
    throw "GitHub token is required. Set GITHUB_TOKEN environment variable, pass -GitHubToken parameter, or ensure .github_token file exists with valid token."
}

if (-not ($Version -match '^\d+\.\d+\.\d+(\.\d+)?$')) {
    throw "Version must be in format x.y.z or x.y.z.w"
}

# Update version information in project files
Write-Host "Updating version information..." -ForegroundColor Yellow
$uiProjectFile = Join-Path $rootDir "BulkEditor.UI\BulkEditor.UI.csproj"

if (Test-Path $uiProjectFile) {
    $content = Get-Content $uiProjectFile -Raw

    # Update AssemblyVersion
    $content = $content -replace '<AssemblyVersion>[^<]*</AssemblyVersion>', "<AssemblyVersion>$Version</AssemblyVersion>"

    # Update FileVersion
    $content = $content -replace '<FileVersion>[^<]*</FileVersion>', "<FileVersion>$Version</FileVersion>"

    # Update Version
    $content = $content -replace '<Version>[^<]*</Version>', "<Version>$Version</Version>"

    Set-Content $uiProjectFile -Value $content -NoNewline
    Write-Host "Updated version information in BulkEditor.UI.csproj to $Version" -ForegroundColor Green
} else {
    Write-Host "Warning: UI project file not found at $uiProjectFile" -ForegroundColor Yellow
}

# Run tests first (only stable tests for deployment)
Write-Host "Running stable tests for deployment..." -ForegroundColor Yellow
dotnet test "$rootDir\BulkEditor.Tests\BulkEditor.Tests.csproj" --configuration Release --filter "FullyQualifiedName~SettingsViewModelTests|FullyQualifiedName~ReplacementServiceTests|FullyQualifiedName~TextReplacementServiceTests|FullyQualifiedName~Core.Entities.DocumentTests|FullyQualifiedName~Core.Entities.HyperlinkTests|FullyQualifiedName~ApplicationServiceTests"
if ($LASTEXITCODE -ne 0) {
    throw "Tests failed. Deployment aborted."
}

# Build installer
Write-Host "Building installer..." -ForegroundColor Yellow
& "$PSScriptRoot\build-installer.ps1" -Configuration Release -Version $Version
if ($LASTEXITCODE -ne 0) {
    throw "Failed to build installer"
}

# Verify installer was created
$installerFile = Join-Path $outputDir "BulkEditor-Setup-$Version.msi"
if (-not (Test-Path $installerFile)) {
    throw "Installer file not found: $installerFile"
}

# Create GitHub release
Write-Host "Creating GitHub release..." -ForegroundColor Yellow

$releaseData = @{
    tag_name = "v$Version"
    target_commitish = "main"
    name = "BulkEditor v$Version"
    body = if ([string]::IsNullOrEmpty($ReleaseNotes)) { "Release v$Version" } else { $ReleaseNotes }
    draft = $false
    prerelease = $Prerelease.IsPresent
} | ConvertTo-Json -Depth 2

$headers = @{
    'Authorization' = "Bearer $GitHubToken"
    'Accept' = 'application/vnd.github.v3+json'
    'Content-Type' = 'application/json'
}

try {
    # Create the release
    $createReleaseUri = "https://api.github.com/repos/ItMeDiaTech/Bulk_Editor_WPF/releases"
    $response = Invoke-RestMethod -Uri $createReleaseUri -Method Post -Body $releaseData -Headers $headers
    $releaseId = $response.id
    $uploadUrl = $response.upload_url -replace '\{\?name,label\}', ''
    Write-Host "Upload URL: $uploadUrl" -ForegroundColor Gray

    Write-Host "GitHub release created with ID: $releaseId" -ForegroundColor Green

    # Upload installer asset
    Write-Host "Uploading installer..." -ForegroundColor Yellow
    $installerBytes = [System.IO.File]::ReadAllBytes($installerFile)
    $installerName = Split-Path $installerFile -Leaf

    $uploadHeaders = @{
        'Authorization' = "Bearer $GitHubToken"
        'Content-Type' = 'application/octet-stream'
    }

    $uploadUri = $uploadUrl + "?name=" + $installerName
    Write-Host "Full upload URI: $uploadUri" -ForegroundColor Gray
    Invoke-RestMethod -Uri $uploadUri -Method Post -Body $installerBytes -Headers $uploadHeaders

    Write-Host "Installer uploaded successfully!" -ForegroundColor Green

    # Upload version.json as additional asset
    Write-Host "Uploading version information..." -ForegroundColor Yellow
    $versionFile = Join-Path $outputDir "version.json"
    if (Test-Path $versionFile) {
        $versionBytes = [System.IO.File]::ReadAllBytes($versionFile)
        $versionUploadUri = $uploadUrl + "?name=version.json"
        Write-Host "Version upload URI: $versionUploadUri" -ForegroundColor Gray
        Invoke-RestMethod -Uri $versionUploadUri -Method Post -Body $versionBytes -Headers $uploadHeaders
        Write-Host "Version information uploaded!" -ForegroundColor Green
    }

    Write-Host "`nRelease Summary:" -ForegroundColor Cyan
    Write-Host "- Version: $Version" -ForegroundColor White
    Write-Host "- Release ID: $releaseId" -ForegroundColor White
    Write-Host "- Installer: $installerName" -ForegroundColor White
    Write-Host "- Prerelease: $($Prerelease.IsPresent)" -ForegroundColor White
    Write-Host "- Release URL: $($response.html_url)" -ForegroundColor White

    Write-Host "`nDeployment completed successfully!" -ForegroundColor Green
}
catch {
    Write-Host "Failed to create GitHub release: $($_.Exception.Message)" -ForegroundColor Red

    # Check if it's a 422 error (release already exists)
    if ($_.Exception.Response.StatusCode -eq 422) {
        Write-Host "Release may already exist. Check GitHub releases page." -ForegroundColor Yellow
    }

    throw
}
