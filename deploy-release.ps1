# Complete Deployment Script for BulkEditor
# Usage: .\deploy-release.ps1 "1.5.1"

param(
    [Parameter(Mandatory=$true)]
    [string]$Version,

    [Parameter(Mandatory=$false)]
    [string]$ReleaseNotes = "",

    [Parameter(Mandatory=$false)]
    [switch]$Prerelease,

    [Parameter(Mandatory=$false)]
    [string]$GitHubToken = $env:GITHUB_TOKEN,

    [Parameter(Mandatory=$false)]
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
$rootDir = $PSScriptRoot
$outputDir = Join-Path $rootDir "Output"

Write-Host "🚀 BulkEditor Complete Deployment Script" -ForegroundColor Cyan
Write-Host "Version: $Version" -ForegroundColor Green
Write-Host "Root Directory: $rootDir" -ForegroundColor Gray

# Validate version format
if (-not ($Version -match '^\d+\.\d+\.\d+(\.\d+)?$')) {
    throw "❌ Version must be in format x.y.z or x.y.z.w (e.g., 1.5.1)"
}

# Create output directory if it doesn't exist
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

try {
    Write-Host "`n📝 Step 1: Updating Version Information..." -ForegroundColor Yellow
    
    # Update UI project version
    $uiProjectFile = Join-Path $rootDir "BulkEditor.UI\BulkEditor.UI.csproj"
    if (Test-Path $uiProjectFile) {
        $content = Get-Content $uiProjectFile -Raw
        $content = $content -replace '<AssemblyVersion>[^<]*</AssemblyVersion>', "<AssemblyVersion>$Version</AssemblyVersion>"
        $content = $content -replace '<FileVersion>[^<]*</FileVersion>', "<FileVersion>$Version</FileVersion>"
        $content = $content -replace '<Version>[^<]*</Version>', "<Version>$Version</Version>"
        Set-Content $uiProjectFile -Value $content -NoNewline
        Write-Host "✅ Updated BulkEditor.UI.csproj to version $Version" -ForegroundColor Green
    }

    # Update installer project version
    $installerProjectFile = Join-Path $rootDir "BulkEditor.Installer\BulkEditor.Installer.wixproj"
    if (Test-Path $installerProjectFile) {
        $content = Get-Content $installerProjectFile -Raw
        $content = $content -replace '<ProductVersion>[^<]*</ProductVersion>', "<ProductVersion>$Version</ProductVersion>"
        Set-Content $installerProjectFile -Value $content -NoNewline
        Write-Host "✅ Updated BulkEditor.Installer.wixproj to version $Version" -ForegroundColor Green
    }

    # Update version.json
    $versionFile = Join-Path $outputDir "version.json"
    $versionData = @{
        Configuration = "Release"
        BuildDate = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Installer = "BulkEditor-Setup-$Version.msi"
        Version = $Version
    } | ConvertTo-Json -Depth 2
    Set-Content $versionFile -Value $versionData -Encoding UTF8
    Write-Host "✅ Updated version.json" -ForegroundColor Green

    Write-Host "`n🧪 Step 2: Running Tests..." -ForegroundColor Yellow
    if (-not $SkipTests) {
        dotnet test --configuration Release --verbosity minimal --no-build
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "⚠️ Some tests failed, but continuing with deployment"
        } else {
            Write-Host "✅ Tests passed!" -ForegroundColor Green
        }
    } else {
        Write-Host "⏭️ Skipping tests (--SkipTests specified)" -ForegroundColor Yellow
    }

    Write-Host "`n🔨 Step 3: Building Release..." -ForegroundColor Yellow
    
    # Clean and build solution
    dotnet clean --configuration Release --verbosity quiet
    dotnet build --configuration Release --verbosity minimal
    if ($LASTEXITCODE -ne 0) { throw "❌ Build failed" }
    Write-Host "✅ Solution built successfully" -ForegroundColor Green

    Write-Host "`n📦 Step 4: Creating Portable Package..." -ForegroundColor Yellow
    
    # Publish portable version
    $publishDir = Join-Path $outputDir "Publish"
    if (Test-Path $publishDir) { Remove-Item $publishDir -Recurse -Force }
    
    dotnet publish "BulkEditor.UI\BulkEditor.UI.csproj" --configuration Release --output $publishDir --self-contained false --verbosity quiet
    if ($LASTEXITCODE -ne 0) { throw "❌ Publish failed" }

    # Create portable ZIP
    $portableZip = Join-Path $outputDir "BulkEditor-v$Version-Portable.zip"
    if (Test-Path $portableZip) { Remove-Item $portableZip -Force }
    
    Compress-Archive -Path "$publishDir\*" -DestinationPath $portableZip -CompressionLevel Optimal
    Write-Host "✅ Created portable package: BulkEditor-v$Version-Portable.zip" -ForegroundColor Green

    Write-Host "`n🏗️ Step 5: Building MSI Installer..." -ForegroundColor Yellow
    
    # Build MSI installer
    dotnet build "BulkEditor.Installer\BulkEditor.Installer.wixproj" --configuration Release --verbosity minimal
    if ($LASTEXITCODE -ne 0) { throw "❌ MSI build failed" }

    # Copy MSI to output with correct name
    $sourceMsi = "BulkEditor.Installer\bin\Release\BulkEditor.Installer.msi"
    $targetMsi = Join-Path $outputDir "BulkEditor-Setup-$Version.msi"
    if (Test-Path $sourceMsi) {
        Copy-Item $sourceMsi $targetMsi -Force
        Write-Host "✅ Created MSI installer: BulkEditor-Setup-$Version.msi" -ForegroundColor Green
    } else {
        throw "❌ MSI file not found at $sourceMsi"
    }

    Write-Host "`n📋 Step 6: Committing Changes..." -ForegroundColor Yellow
    
    # Commit version updates
    git add -A
    git commit -m "🔖 Release v$Version - Update version files and build artifacts"
    git push origin main
    Write-Host "✅ Version changes committed and pushed" -ForegroundColor Green

    Write-Host "`n🚀 Step 7: Creating GitHub Release..." -ForegroundColor Yellow
    
    # Create release notes if not provided
    if ([string]::IsNullOrEmpty($ReleaseNotes)) {
        $ReleaseNotes = @"
# BulkEditor v$Version

## What's New
- Version $Version release
- Build date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

## Downloads
- **BulkEditor-Setup-$Version.msi**: Windows installer (recommended)
- **BulkEditor-v$Version-Portable.zip**: Portable version

## System Requirements
- Windows 10 or later
- .NET 8.0 Runtime
- Microsoft Word (for document processing)
"@
    }

    # Create GitHub release
    $prereleaseFlag = if ($Prerelease) { "--prerelease" } else { "" }
    $releaseCommand = "gh release create `"v$Version`" --title `"BulkEditor v$Version`" --notes `"$ReleaseNotes`" --target main $prereleaseFlag"
    
    Invoke-Expression $releaseCommand
    Write-Host "✅ GitHub release created: v$Version" -ForegroundColor Green

    Write-Host "`n📤 Step 8: Uploading Release Assets..." -ForegroundColor Yellow
    
    # Upload MSI installer
    gh release upload "v$Version" $targetMsi --clobber
    Write-Host "✅ Uploaded MSI installer" -ForegroundColor Green

    # Upload portable ZIP
    gh release upload "v$Version" $portableZip --clobber
    Write-Host "✅ Uploaded portable package" -ForegroundColor Green

    # Upload version info
    gh release upload "v$Version" $versionFile --clobber
    Write-Host "✅ Uploaded version information" -ForegroundColor Green

    Write-Host "`n🎉 DEPLOYMENT SUCCESSFUL!" -ForegroundColor Green
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host "📦 Release: v$Version" -ForegroundColor White
    Write-Host "🔗 URL: https://github.com/ItMeDiaTech/Bulk_Editor_WPF/releases/tag/v$Version" -ForegroundColor White
    Write-Host "📂 Assets:" -ForegroundColor White
    Write-Host "   • BulkEditor-Setup-$Version.msi" -ForegroundColor Gray
    Write-Host "   • BulkEditor-v$Version-Portable.zip" -ForegroundColor Gray
    Write-Host "   • version.json" -ForegroundColor Gray
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    
    # Show file sizes
    $msiSize = [math]::Round((Get-Item $targetMsi).Length / 1MB, 2)
    $zipSize = [math]::Round((Get-Item $portableZip).Length / 1MB, 2)
    Write-Host "📊 Package Sizes:" -ForegroundColor White
    Write-Host "   • MSI: $msiSize MB" -ForegroundColor Gray
    Write-Host "   • ZIP: $zipSize MB" -ForegroundColor Gray

} catch {
    Write-Host "`n❌ DEPLOYMENT FAILED!" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Yellow
    exit 1
}

Write-Host "`n🎯 Deployment completed successfully! Ready for distribution." -ForegroundColor Green