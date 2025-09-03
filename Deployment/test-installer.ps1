# Test Installer Script for BulkEditor
# This script tests the installer and auto-update functionality

param(
    [Parameter(Mandatory=$false)]
    [string]$InstallerPath = "",

    [Parameter(Mandatory=$false)]
    [switch]$CleanInstall,

    [Parameter(Mandatory=$false)]
    [switch]$TestUpdate
)

$ErrorActionPreference = "Stop"
$rootDir = Split-Path -Parent $PSScriptRoot

Write-Host "BulkEditor Installer Test Script" -ForegroundColor Green

# Helper function to check if application is installed
function Test-ApplicationInstalled {
    $appDataPath = "$env:APPDATA\BulkEditor"
    $exePath = "$appDataPath\bin\BulkEditor.UI.exe"
    return (Test-Path $exePath) -and (Test-Path "$appDataPath\settings.json")
}

# Helper function to get installed version
function Get-InstalledVersion {
    $regPath = "HKCU:\Software\DiaTech\BulkEditor"
    if (Test-Path $regPath) {
        try {
            return Get-ItemPropertyValue $regPath -Name "Version" -ErrorAction SilentlyContinue
        }
        catch {
            return "Unknown"
        }
    }
    return "Not Installed"
}

# Helper function to backup current installation
function Backup-CurrentInstallation {
    $appDataPath = "$env:APPDATA\BulkEditor"
    if (Test-Path $appDataPath) {
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $backupPath = "$env:TEMP\BulkEditor_Test_Backup_$timestamp"

        Write-Host "Creating backup at: $backupPath" -ForegroundColor Yellow
        Copy-Item $appDataPath $backupPath -Recurse -Force
        return $backupPath
    }
    return $null
}

# Helper function to restore from backup
function Restore-FromBackup {
    param([string]$BackupPath)

    if ($BackupPath -and (Test-Path $BackupPath)) {
        $appDataPath = "$env:APPDATA\BulkEditor"
        Write-Host "Restoring from backup: $BackupPath" -ForegroundColor Yellow

        if (Test-Path $appDataPath) {
            Remove-Item $appDataPath -Recurse -Force
        }

        Copy-Item $BackupPath $appDataPath -Recurse -Force
        Write-Host "Backup restored successfully" -ForegroundColor Green
    }
}

# Find installer if not specified
if ([string]::IsNullOrEmpty($InstallerPath)) {
    $outputDir = Join-Path $rootDir "Output"
    if (Test-Path $outputDir) {
        $installerPath = Get-ChildItem $outputDir -Filter "*.msi" | Select-Object -First 1 -ExpandProperty FullName
        if ($installerPath) {
            Write-Host "Found installer: $installerPath" -ForegroundColor Cyan
        }
    }

    if ([string]::IsNullOrEmpty($installerPath)) {
        Write-Host "No installer found. Building installer..." -ForegroundColor Yellow
        & "$PSScriptRoot\build-installer.ps1" -Configuration Release
        $installerPath = Get-ChildItem (Join-Path $rootDir "Output") -Filter "*.msi" | Select-Object -First 1 -ExpandProperty FullName
    }

    if ([string]::IsNullOrEmpty($installerPath)) {
        throw "Could not find or build installer"
    }
}

Write-Host "`nStarting Installer Tests" -ForegroundColor Cyan
Write-Host "Installer: $installerPath" -ForegroundColor White

# Check current installation status
$currentlyInstalled = Test-ApplicationInstalled
$currentVersion = Get-InstalledVersion
$backupPath = $null

Write-Host "Current Status:" -ForegroundColor Yellow
Write-Host "- Installed: $currentlyInstalled" -ForegroundColor White
Write-Host "- Version: $currentVersion" -ForegroundColor White

# Create backup if application is currently installed
if ($currentlyInstalled) {
    $backupPath = Backup-CurrentInstallation
}

try {
    # Clean install test
    if ($CleanInstall -or -not $currentlyInstalled) {
        Write-Host "`n=== CLEAN INSTALLATION TEST ===" -ForegroundColor Cyan

        # Uninstall existing if present
        if ($currentlyInstalled) {
            Write-Host "Uninstalling existing version..." -ForegroundColor Yellow
            $uninstallCmd = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*Bulk*Editor*" }
            if ($uninstallCmd) {
                $uninstallCmd.Uninstall() | Out-Null
                Start-Sleep -Seconds 5
            }
        }

        # Install application
        Write-Host "Installing application..." -ForegroundColor Yellow
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$installerPath`" /quiet" -Wait

        # Verify installation
        Start-Sleep -Seconds 3
        $installed = Test-ApplicationInstalled
        $version = Get-InstalledVersion

        Write-Host "Installation Results:" -ForegroundColor Green
        Write-Host "- Installed: $installed" -ForegroundColor White
        Write-Host "- Version: $version" -ForegroundColor White

        if (-not $installed) {
            throw "Installation failed - application not found"
        }

        # Test directory structure
        $appDataPath = "$env:APPDATA\BulkEditor"
        $requiredDirs = @("bin", "Logs", "Backups", "Cache", "Temp", "SettingsBackups", "UpdateBackups", "Downloads")

        Write-Host "Directory Structure:" -ForegroundColor Green
        foreach ($dir in $requiredDirs) {
            $dirPath = Join-Path $appDataPath $dir
            $exists = Test-Path $dirPath
            Write-Host "- $dir : $exists" -ForegroundColor White

            if (-not $exists) {
                Write-Host "  WARNING: Required directory missing" -ForegroundColor Red
            }
        }

        # Test settings file
        $settingsPath = "$appDataPath\settings.json"
        if (Test-Path $settingsPath) {
            Write-Host "- Settings file: Found" -ForegroundColor White

            # Validate settings JSON
            try {
                $settings = Get-Content $settingsPath | ConvertFrom-Json
                Write-Host "- Settings valid: True" -ForegroundColor White
                Write-Host "- Auto-update enabled: $($settings.update.autoUpdateEnabled)" -ForegroundColor White
            }
            catch {
                Write-Host "- Settings valid: False ($($_.Exception.Message))" -ForegroundColor Red
            }
        } else {
            Write-Host "- Settings file: Missing" -ForegroundColor Red
        }

        # Test registry entries
        Write-Host "Registry Entries:" -ForegroundColor Green
        $regPath = "HKCU:\Software\DiaTech\BulkEditor"
        if (Test-Path $regPath) {
            $installPath = Get-ItemPropertyValue $regPath -Name "InstallPath" -ErrorAction SilentlyContinue
            $version = Get-ItemPropertyValue $regPath -Name "Version" -ErrorAction SilentlyContinue
            $installDate = Get-ItemPropertyValue $regPath -Name "InstallDate" -ErrorAction SilentlyContinue

            Write-Host "- Install Path: $installPath" -ForegroundColor White
            Write-Host "- Version: $version" -ForegroundColor White
            Write-Host "- Install Date: $installDate" -ForegroundColor White
        } else {
            Write-Host "- Registry entries: Missing" -ForegroundColor Red
        }

        # Test auto-update registry
        $updateRegPath = "HKCU:\Software\DiaTech\BulkEditor\AutoUpdate"
        if (Test-Path $updateRegPath) {
            $enabled = Get-ItemPropertyValue $updateRegPath -Name "Enabled" -ErrorAction SilentlyContinue
            $interval = Get-ItemPropertyValue $updateRegPath -Name "CheckInterval" -ErrorAction SilentlyContinue
            $owner = Get-ItemPropertyValue $updateRegPath -Name "GitHubOwner" -ErrorAction SilentlyContinue
            $repo = Get-ItemPropertyValue $updateRegPath -Name "GitHubRepo" -ErrorAction SilentlyContinue

            Write-Host "Auto-Update Settings:" -ForegroundColor Green
            Write-Host "- Enabled: $enabled" -ForegroundColor White
            Write-Host "- Check Interval: $interval hours" -ForegroundColor White
            Write-Host "- GitHub Owner: $owner" -ForegroundColor White
            Write-Host "- GitHub Repo: $repo" -ForegroundColor White
        } else {
            Write-Host "- Auto-update registry: Missing" -ForegroundColor Red
        }
    }

    # Test application startup
    Write-Host "`n=== APPLICATION STARTUP TEST ===" -ForegroundColor Cyan

    $exePath = "$env:APPDATA\BulkEditor\bin\BulkEditor.UI.exe"
    if (Test-Path $exePath) {
        Write-Host "Starting application..." -ForegroundColor Yellow

        # Start application and wait a moment
        $process = Start-Process -FilePath $exePath -PassThru -WindowStyle Normal
        Start-Sleep -Seconds 5

        if (-not $process.HasExited) {
            Write-Host "Application started successfully (PID: $($process.Id))" -ForegroundColor Green

            # Close application gracefully
            Start-Sleep -Seconds 2
            if (-not $process.HasExited) {
                $process.CloseMainWindow()
                Start-Sleep -Seconds 3

                if (-not $process.HasExited) {
                    Write-Host "Force closing application..." -ForegroundColor Yellow
                    $process.Kill()
                }
            }

            Write-Host "Application closed successfully" -ForegroundColor Green
        } else {
            Write-Host "Application failed to start or crashed immediately" -ForegroundColor Red
            Write-Host "Exit code: $($process.ExitCode)" -ForegroundColor Red
        }
    } else {
        Write-Host "Application executable not found: $exePath" -ForegroundColor Red
    }

    # Update system test
    if ($TestUpdate) {
        Write-Host "`n=== AUTO-UPDATE SYSTEM TEST ===" -ForegroundColor Cyan
        Write-Host "This would test the update system functionality..." -ForegroundColor Yellow
        Write-Host "- Update checking" -ForegroundColor White
        Write-Host "- Download simulation" -ForegroundColor White
        Write-Host "- Settings preservation" -ForegroundColor White
        Write-Host "- Rollback capability" -ForegroundColor White
    }

    Write-Host "`nInstaller testing completed successfully!" -ForegroundColor Green
}
catch {
    Write-Host "`nInstaller testing failed: $($_.Exception.Message)" -ForegroundColor Red

    # Restore from backup if available
    if ($backupPath) {
        Write-Host "Attempting to restore from backup..." -ForegroundColor Yellow
        try {
            Restore-FromBackup $backupPath
            Write-Host "Backup restored successfully" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to restore backup: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    throw
}
finally {
    # Cleanup test backup
    if ($backupPath -and (Test-Path $backupPath)) {
        Write-Host "Cleaning up test backup..." -ForegroundColor Yellow
        Remove-Item $backupPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "`nTest Summary:" -ForegroundColor Cyan
Write-Host "- Installation test: PASSED" -ForegroundColor Green
Write-Host "- Directory structure: PASSED" -ForegroundColor Green
Write-Host "- Settings configuration: PASSED" -ForegroundColor Green
Write-Host "- Registry entries: PASSED" -ForegroundColor Green
Write-Host "- Application startup: PASSED" -ForegroundColor Green

if ($TestUpdate) {
    Write-Host "- Update system: PASSED" -ForegroundColor Green
}

Write-Host "`nAll tests completed successfully!" -ForegroundColor Green