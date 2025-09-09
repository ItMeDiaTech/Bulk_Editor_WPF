# 🚀 BulkEditor Deployment Guide

This guide explains how to use the automated deployment scripts to build, package, and release new versions of BulkEditor.

## 📋 Prerequisites

### Required Tools
- **.NET 8.0 SDK** (for building)
- **Git** (for version control)
- **GitHub CLI (`gh`)** (for creating releases)
- **PowerShell** (Windows) or **Bash** (Cross-platform)

### GitHub Setup
1. Install GitHub CLI: `winget install GitHub.cli` (Windows) or `brew install gh` (macOS)
2. Authenticate: `gh auth login`
3. Ensure you have write access to the repository

## 🎯 Quick Start

### PowerShell (Windows - Recommended)
```powershell
# Deploy new version with automatic everything
.\deploy-release.ps1 "1.5.2"

# Deploy with custom release notes
.\deploy-release.ps1 "1.5.2" -ReleaseNotes "Bug fixes and improvements"

# Deploy as prerelease
.\deploy-release.ps1 "1.6.0-beta.1" -Prerelease

# Skip tests (faster deployment)
.\deploy-release.ps1 "1.5.2" -SkipTests
```

### Bash (Cross-platform)
```bash
# Deploy new version
./deploy-release.sh "1.5.2"

# Deploy with custom release notes
./deploy-release.sh "1.5.2" "Bug fixes and improvements"

# Deploy as prerelease
./deploy-release.sh "1.6.0-beta.1" "" "true"
```

## 📦 What the Script Does

### Automated Steps
1. **📝 Version Updates**
   - Updates `BulkEditor.UI.csproj` version numbers
   - Updates `BulkEditor.Installer.wixproj` product version
   - Updates `Output/version.json` with build info

2. **🔨 Building**
   - Cleans previous builds
   - Builds entire solution in Release configuration
   - Runs tests (unless `--SkipTests` specified)

3. **📦 Packaging**
   - Creates portable ZIP package (`BulkEditor-v{version}-Portable.zip`)
   - Builds MSI installer (`BulkEditor-Setup-{version}.msi`)
   - Generates version metadata file

4. **🚀 Publishing**
   - Commits version changes to Git
   - Creates GitHub release with tag `v{version}`
   - Uploads all three files (MSI, ZIP, version.json)
   - Provides download links and file sizes

## 📂 Output Files

After successful deployment, you'll find these files in the `Output/` directory:

### 📥 Release Assets (uploaded to GitHub)
- **`BulkEditor-Setup-{version}.msi`** - Windows installer (recommended for users)
- **`BulkEditor-v{version}-Portable.zip`** - Portable version (no installation required)
- **`version.json`** - Build metadata for auto-updater

### 🗂️ Local Files
- **`Output/Publish/`** - Unpacked portable version
- **`BulkEditor.Installer/bin/Release/`** - Built installer files

## 🎨 Examples

### 🔄 Patch Release
```powershell
.\deploy-release.ps1 "1.5.3"
```
- Incremental bug fixes
- Automatic release notes
- Full deployment pipeline

### 🆕 Minor Release
```powershell
$releaseNotes = @"
# BulkEditor v1.6.0 - New Features

## ✨ What's New
- Enhanced hyperlink processing engine
- Improved error handling and logging
- New settings for advanced users

## 🐛 Bug Fixes
- Fixed memory leaks in document processing
- Resolved UI freezing issues
- Better handling of large documents

## 🔄 Breaking Changes
None - fully backward compatible
"@

.\deploy-release.ps1 "1.6.0" -ReleaseNotes $releaseNotes
```

### 🧪 Beta Release
```powershell
.\deploy-release.ps1 "2.0.0-beta.1" -Prerelease -ReleaseNotes "Beta release with experimental features"
```

## ⚡ Quick Deploy Commands

### Most Common Use Cases
```powershell
# Just bump patch version and deploy
.\deploy-release.ps1 "1.5.4"

# Deploy with brief notes
.\deploy-release.ps1 "1.5.4" -ReleaseNotes "Critical security fixes"

# Fast deployment (skip tests)
.\deploy-release.ps1 "1.5.4" -SkipTests
```

## 🔍 Troubleshooting

### Common Issues

#### ❌ "Version must be in format x.y.z"
- Ensure version follows semantic versioning: `1.2.3` or `1.2.3.4`
- Examples: ✅ `1.5.2` ✅ `2.0.0-beta.1` ❌ `v1.5.2` ❌ `1.5`

#### ❌ "gh: command not found"
- Install GitHub CLI: `winget install GitHub.cli`
- Authenticate: `gh auth login`

#### ❌ "Build failed"
- Run `dotnet build` manually to see detailed errors
- Ensure all dependencies are installed
- Check that .NET 8.0 SDK is available

#### ❌ "Tests failed"
- Use `-SkipTests` flag to deploy anyway
- Or fix failing tests before deployment
- Tests are run in Release configuration

### Manual Recovery
If deployment fails partway through:

1. **Check Git status**: `git status` (may need to commit/reset changes)
2. **Check GitHub releases**: Visit GitHub to see if release was created
3. **Clean output**: Remove `Output/Publish` and rebuild
4. **Re-run**: Use the same version number (script handles duplicates)

## 🎯 Best Practices

### Version Numbering
- **Patch (1.5.1 → 1.5.2)**: Bug fixes, small improvements
- **Minor (1.5.x → 1.6.0)**: New features, non-breaking changes  
- **Major (1.x.x → 2.0.0)**: Breaking changes, major redesigns
- **Pre-release (2.0.0-beta.1)**: Testing versions, use `--Prerelease`

### Release Notes
- Keep them concise but informative
- Use markdown formatting for better readability
- Include sections: What's New, Bug Fixes, Breaking Changes
- Mention system requirements if they changed

### Testing
- Always test locally before deployment
- Use `--SkipTests` only for hotfixes or when tests are known to be flaky
- Consider running full test suite manually for major releases

## 🔗 Related Files

- **`deploy-release.ps1`** - Main PowerShell deployment script
- **`deploy-release.sh`** - Bash version for cross-platform use
- **`Deployment/deploy-release.ps1`** - Legacy script (can be removed)
- **`Deployment/deploy.ps1`** - Legacy build-only script (can be removed)

---

**💡 Pro Tip**: Bookmark this guide and the deployment command for quick access!

For issues or improvements to the deployment scripts, please create an issue in the GitHub repository.