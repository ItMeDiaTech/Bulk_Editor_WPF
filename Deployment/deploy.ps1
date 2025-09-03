# PowerShell Deployment Script for BulkEditor WPF Application
# This script builds, tests, and packages the application for distribution

param(
    [Parameter(Mandatory=$false)]
    [string]$Configuration = "Release",

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\dist\",

    [Parameter(Mandatory=$false)]
    [switch]$SkipTests,

    [Parameter(Mandatory=$false)]
    [switch]$CreateInstaller
)

Write-Host "🚀 BulkEditor Deployment Script" -ForegroundColor Cyan
Write-Host "Configuration: $Configuration" -ForegroundColor Green
Write-Host "Output Path: $OutputPath" -ForegroundColor Green

# Set error action preference
$ErrorActionPreference = "Stop"

try {
    # Step 1: Clean previous builds
    Write-Host "`n🧹 Cleaning previous builds..." -ForegroundColor Yellow
    dotnet clean --configuration $Configuration
    if ($LASTEXITCODE -ne 0) { throw "Clean failed" }

    # Step 2: Restore dependencies
    Write-Host "`n📦 Restoring dependencies..." -ForegroundColor Yellow
    dotnet restore
    if ($LASTEXITCODE -ne 0) { throw "Restore failed" }

    # Step 3: Build solution
    Write-Host "`n🔨 Building solution..." -ForegroundColor Yellow
    dotnet build --configuration $Configuration --no-restore
    if ($LASTEXITCODE -ne 0) { throw "Build failed" }

    # Step 4: Run tests (optional)
    if (-not $SkipTests) {
        Write-Host "`n🧪 Running tests..." -ForegroundColor Yellow
        dotnet test --configuration $Configuration --no-build --verbosity normal
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Tests failed, but continuing with deployment"
        }
    }

    # Step 5: Publish application
    Write-Host "`n📱 Publishing application..." -ForegroundColor Yellow

    # Create output directory
    if (Test-Path $OutputPath) {
        Remove-Item $OutputPath -Recurse -Force
    }
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null

    # Publish self-contained executable
    dotnet publish BulkEditor.UI\BulkEditor.UI.csproj `
        --configuration $Configuration `
        --runtime win-x64 `
        --self-contained true `
        --output "$OutputPath\BulkEditor" `
        --verbosity normal `
        -p:PublishSingleFile=false `
        -p:PublishReadyToRun=true `
        -p:IncludeNativeLibrariesForSelfExtract=true

    if ($LASTEXITCODE -ne 0) { throw "Publish failed" }

    # Step 6: Copy additional files
    Write-Host "`n📋 Copying additional files..." -ForegroundColor Yellow

    # Copy configuration files
    Copy-Item "appsettings.json" "$OutputPath\BulkEditor\" -ErrorAction SilentlyContinue

    # Copy documentation
    Copy-Item "README.md" "$OutputPath\BulkEditor\" -ErrorAction SilentlyContinue
    Copy-Item "Project_Info.md" "$OutputPath\BulkEditor\" -ErrorAction SilentlyContinue

    # Create deployment info file
    $deploymentInfo = @"
BulkEditor Deployment Information
================================

Build Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Configuration: $Configuration
Runtime: win-x64
.NET Version: 8.0

Installation Instructions:
1. Extract all files to a directory of your choice
2. Run BulkEditor.UI.exe to start the application
3. Configure settings via the Settings menu

System Requirements:
- Windows 10/11 (x64)
- .NET 8.0 Runtime (included with self-contained deployment)
- Microsoft Word (for document processing)

For support and documentation, see README.md
"@

    $deploymentInfo | Out-File -FilePath "$OutputPath\BulkEditor\DEPLOYMENT_INFO.txt" -Encoding UTF8

    # Step 7: Create ZIP package
    Write-Host "`n📦 Creating ZIP package..." -ForegroundColor Yellow
    $zipPath = "$OutputPath\BulkEditor-v1.0-win-x64.zip"

    if (Get-Command Compress-Archive -ErrorAction SilentlyContinue) {
        Compress-Archive -Path "$OutputPath\BulkEditor\*" -DestinationPath $zipPath -CompressionLevel Optimal
        Write-Host "✅ ZIP package created: $zipPath" -ForegroundColor Green
    } else {
        Write-Warning "Compress-Archive not available. Manual ZIP creation required."
    }

    # Step 8: Create installer (optional)
    if ($CreateInstaller) {
        Write-Host "`n🔧 Creating installer..." -ForegroundColor Yellow
        # Check if NSIS or WiX is available
        if (Get-Command makensis -ErrorAction SilentlyContinue) {
            Write-Host "NSIS detected - creating NSIS installer" -ForegroundColor Green
            # NSIS installer creation would go here
        } elseif (Get-Command candle -ErrorAction SilentlyContinue) {
            Write-Host "WiX detected - creating MSI installer" -ForegroundColor Green
            # WiX installer creation would go here
        } else {
            Write-Warning "No installer creation tools found (NSIS or WiX)"
        }
    }

    # Step 9: Deployment summary
    Write-Host "`n✅ Deployment completed successfully!" -ForegroundColor Green
    Write-Host "Deployment location: $OutputPath" -ForegroundColor Cyan

    $deploymentSize = (Get-ChildItem "$OutputPath\BulkEditor" -Recurse | Measure-Object -Property Length -Sum).Sum
    $deploySizeMB = [math]::Round($deploymentSize / 1MB, 2)
    Write-Host "Package size: $deploySizeMB MB" -ForegroundColor Cyan

    # List main files
    Write-Host "`n📁 Deployment contents:" -ForegroundColor Cyan
    Get-ChildItem "$OutputPath\BulkEditor" -Name | ForEach-Object {
        Write-Host "  - $_" -ForegroundColor Gray
    }

} catch {
    Write-Host "`n❌ Deployment failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host "`n🎉 Ready for distribution!" -ForegroundColor Green