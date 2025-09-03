# BulkEditor Installer

This directory contains the WiX installer project for BulkEditor, which creates a standalone installer that installs the application to the user's AppData directory and includes auto-update functionality.

## Features

- **AppData Installation**: Installs to `%APPDATA%\BulkEditor` for user-specific installations
- **Automatic Directory Structure**: Creates all necessary directories (Logs, Backups, Cache, etc.)
- **Desktop & Start Menu Shortcuts**: Optional shortcuts for easy access
- **Auto-Update Configuration**: Sets up auto-update registry entries
- **Settings Migration**: Migrates existing settings from old locations
- **Uninstall Support**: Clean uninstallation with settings preservation option

## Installation Directory Structure

```
%APPDATA%\BulkEditor\
├── bin\                    # Application executables and libraries
├── Logs\                   # Application logs
├── Backups\                # Document backups
├── Cache\                  # Application cache
├── Temp\                   # Temporary files
├── SettingsBackups\        # Settings backup files
├── UpdateBackups\          # Update backup files
├── Downloads\              # Update downloads
└── settings.json           # User settings
```

## Building the Installer

### Prerequisites

1. Install WiX Toolset v4.0+:

   ```bash
   dotnet tool install --global wix
   ```

2. Install .NET 6.0+ SDK

### Manual Build

```bash
# From the installer directory
dotnet build --configuration Release
```

### Automated Build

Use the provided PowerShell script:

```powershell
# Build with default version
.\Deployment\build-installer.ps1

# Build specific version
.\Deployment\build-installer.ps1 -Version "1.2.3"

# Skip application build (if already built)
.\Deployment\build-installer.ps1 -SkipBuild
```

## Installer Features

### Installation Options

- **Installation Directory**: Fixed to `%APPDATA%\BulkEditor` for user-specific installations
- **Desktop Shortcut**: Optional (default: enabled)
- **Start Menu Shortcuts**: Optional (default: enabled)
- **Auto-Update**: Enabled by default

### Registry Entries

The installer creates the following registry entries:

```
HKCU\Software\DiaTech\BulkEditor\
├── InstallPath            # Installation directory
├── Version               # Installed version
└── InstallDate          # Installation date

HKCU\Software\DiaTech\BulkEditor\AutoUpdate\
├── Enabled              # Auto-update enabled (1/0)
├── CheckInterval        # Check interval in hours
├── LastCheck           # Last update check timestamp
├── GitHubOwner         # GitHub repository owner
└── GitHubRepo          # GitHub repository name
```

### Silent Installation

The installer supports silent installation for automated deployments:

```bash
# Silent install
BulkEditor-Setup-1.0.0.msi /S

# Silent install to specific directory
BulkEditor-Setup-1.0.0.msi /S /D="C:\Custom\Path"
```

## Auto-Update System

### Configuration

Auto-update settings are configured through:

1. **Registry entries** (set during installation)
2. **Application settings** (user-configurable)
3. **Command-line parameters** (for testing)

### Update Process

1. **Check for Updates**: Queries GitHub API for latest release
2. **Download**: Downloads installer to `%APPDATA%\BulkEditor\Downloads`
3. **Backup**: Creates backup of current settings and installation
4. **Install**: Runs silent installer and restarts application
5. **Verify**: Confirms successful update and restores settings if needed

### Update Flow

```
Application Startup → Check Last Update Time →
If Due → Check GitHub API → If Update Available →
Notify User → User Confirms → Download → Install → Restart
```

## Deployment Workflow

### 1. Build and Test

```powershell
# Build the installer
.\Deployment\build-installer.ps1 -Version "1.0.0"

# Test the installer locally
.\Output\BulkEditor-Setup-1.0.0.msi
```

### 2. Deploy to GitHub

```powershell
# Deploy release with release notes
.\Deployment\deploy-release.ps1 -Version "1.0.0" -ReleaseNotes "Initial release with auto-update support"

# Deploy prerelease
.\Deployment\deploy-release.ps1 -Version "1.0.0-beta.1" -Prerelease
```

## Troubleshooting

### Common Issues

1. **WiX Build Errors**:

   - Ensure WiX Toolset v4.0+ is installed
   - Check that all referenced files exist
   - Verify project references are correct

2. **File Path Issues**:

   - Use relative paths in the WiX project
   - Ensure output directories exist
   - Check file permissions

3. **Registry Issues**:
   - Run installer as administrator if needed
   - Check Windows permission policies
   - Verify registry path syntax

### Log Files

Check the following locations for troubleshooting:

- **Build Logs**: `BulkEditor.Installer\bin\Release\`
- **MSI Logs**: `%TEMP%\MSI*.log` (enable with `/l*v logfile.log`)
- **Application Logs**: `%APPDATA%\BulkEditor\Logs\`

## Security Considerations

- Installer is unsigned by default - consider code signing for production
- Auto-update checks GitHub over HTTPS
- Settings and logs stored in user profile (secure by default)
- No elevation required for installation
- Update downloads verified by size and basic integrity checks

## Development Notes

- The installer uses per-user installation scope
- Settings are automatically migrated from old locations
- Update system preserves user settings across updates
- Rollback capability for failed updates
- Cleanup of old backups and temporary files
