#!/bin/bash
# Complete Deployment Script for BulkEditor (Bash version)
# Usage: ./deploy-release.sh "1.5.1"

set -e  # Exit on any error

VERSION="$1"
RELEASE_NOTES="${2:-}"
PRERELEASE="${3:-false}"

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

if [[ -z "$VERSION" ]]; then
    echo -e "${RED}❌ Usage: ./deploy-release.sh \"x.y.z\" [release_notes] [prerelease]${NC}"
    echo -e "${YELLOW}Example: ./deploy-release.sh \"1.5.1\"${NC}"
    exit 1
fi

# Validate version format
if [[ ! "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+(\.[0-9]+)?$ ]]; then
    echo -e "${RED}❌ Version must be in format x.y.z or x.y.z.w (e.g., 1.5.1)${NC}"
    exit 1
fi

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUTPUT_DIR="$ROOT_DIR/Output"

echo -e "${CYAN}🚀 BulkEditor Complete Deployment Script${NC}"
echo -e "${GREEN}Version: $VERSION${NC}"
echo -e "${BLUE}Root Directory: $ROOT_DIR${NC}"

# Create output directory if it doesn't exist
mkdir -p "$OUTPUT_DIR"

echo -e "\n${YELLOW}📝 Step 1: Updating Version Information...${NC}"

# Update UI project version
UI_PROJECT_FILE="$ROOT_DIR/BulkEditor.UI/BulkEditor.UI.csproj"
if [[ -f "$UI_PROJECT_FILE" ]]; then
    sed -i.bak "s|<AssemblyVersion>[^<]*</AssemblyVersion>|<AssemblyVersion>$VERSION</AssemblyVersion>|g" "$UI_PROJECT_FILE"
    sed -i.bak "s|<FileVersion>[^<]*</FileVersion>|<FileVersion>$VERSION</FileVersion>|g" "$UI_PROJECT_FILE"
    sed -i.bak "s|<Version>[^<]*</Version>|<Version>$VERSION</Version>|g" "$UI_PROJECT_FILE"
    rm "$UI_PROJECT_FILE.bak" 2>/dev/null || true
    echo -e "${GREEN}✅ Updated BulkEditor.UI.csproj to version $VERSION${NC}"
fi

# Update installer project version
INSTALLER_PROJECT_FILE="$ROOT_DIR/BulkEditor.Installer/BulkEditor.Installer.wixproj"
if [[ -f "$INSTALLER_PROJECT_FILE" ]]; then
    sed -i.bak "s|<ProductVersion>[^<]*</ProductVersion>|<ProductVersion>$VERSION</ProductVersion>|g" "$INSTALLER_PROJECT_FILE"
    rm "$INSTALLER_PROJECT_FILE.bak" 2>/dev/null || true
    echo -e "${GREEN}✅ Updated BulkEditor.Installer.wixproj to version $VERSION${NC}"
fi

# Update version.json
VERSION_FILE="$OUTPUT_DIR/version.json"
BUILD_DATE=$(date "+%Y-%m-%d %H:%M:%S")
cat > "$VERSION_FILE" << EOF
{
    "Configuration": "Release",
    "BuildDate": "$BUILD_DATE",
    "Installer": "BulkEditor-Setup-$VERSION.msi",
    "Version": "$VERSION"
}
EOF
echo -e "${GREEN}✅ Updated version.json${NC}"

echo -e "\n${YELLOW}🔨 Step 2: Building Release...${NC}"

# Clean and build solution
dotnet clean --configuration Release --verbosity quiet
if ! dotnet build --configuration Release --verbosity minimal; then
    echo -e "${RED}❌ Build failed${NC}"
    exit 1
fi
echo -e "${GREEN}✅ Solution built successfully${NC}"

echo -e "\n${YELLOW}📦 Step 3: Creating Portable Package...${NC}"

# Publish portable version
PUBLISH_DIR="$OUTPUT_DIR/Publish"
rm -rf "$PUBLISH_DIR"

if ! dotnet publish "BulkEditor.UI/BulkEditor.UI.csproj" --configuration Release --output "$PUBLISH_DIR" --self-contained false --verbosity quiet; then
    echo -e "${RED}❌ Publish failed${NC}"
    exit 1
fi

# Create portable ZIP
PORTABLE_ZIP="$OUTPUT_DIR/BulkEditor-v$VERSION-Portable.zip"
rm -f "$PORTABLE_ZIP"

cd "$PUBLISH_DIR" && tar -czf "$PORTABLE_ZIP" . && cd "$ROOT_DIR"
echo -e "${GREEN}✅ Created portable package: BulkEditor-v$VERSION-Portable.zip${NC}"

echo -e "\n${YELLOW}🏗️ Step 4: Building MSI Installer...${NC}"

# Build MSI installer
if ! dotnet build "BulkEditor.Installer/BulkEditor.Installer.wixproj" --configuration Release --verbosity minimal; then
    echo -e "${RED}❌ MSI build failed${NC}"
    exit 1
fi

# Copy MSI to output with correct name
SOURCE_MSI="BulkEditor.Installer/bin/Release/BulkEditor.Installer.msi"
TARGET_MSI="$OUTPUT_DIR/BulkEditor-Setup-$VERSION.msi"
if [[ -f "$SOURCE_MSI" ]]; then
    cp "$SOURCE_MSI" "$TARGET_MSI"
    echo -e "${GREEN}✅ Created MSI installer: BulkEditor-Setup-$VERSION.msi${NC}"
else
    echo -e "${RED}❌ MSI file not found at $SOURCE_MSI${NC}"
    exit 1
fi

echo -e "\n${YELLOW}📋 Step 5: Committing Changes...${NC}"

# Commit version updates
git add -A
git commit -m "🔖 Release v$VERSION - Update version files and build artifacts"
git push origin main
echo -e "${GREEN}✅ Version changes committed and pushed${NC}"

echo -e "\n${YELLOW}🚀 Step 6: Creating GitHub Release...${NC}"

# Create release notes if not provided
if [[ -z "$RELEASE_NOTES" ]]; then
    RELEASE_NOTES="# BulkEditor v$VERSION

## What's New
- Version $VERSION release
- Build date: $BUILD_DATE

## Downloads
- **BulkEditor-Setup-$VERSION.msi**: Windows installer (recommended)
- **BulkEditor-v$VERSION-Portable.zip**: Portable version

## System Requirements
- Windows 10 or later
- .NET 8.0 Runtime
- Microsoft Word (for document processing)"
fi

# Create GitHub release
PRERELEASE_FLAG=""
if [[ "$PRERELEASE" == "true" ]]; then
    PRERELEASE_FLAG="--prerelease"
fi

gh release create "v$VERSION" --title "BulkEditor v$VERSION" --notes "$RELEASE_NOTES" --target main $PRERELEASE_FLAG
echo -e "${GREEN}✅ GitHub release created: v$VERSION${NC}"

echo -e "\n${YELLOW}📤 Step 7: Uploading Release Assets...${NC}"

# Upload MSI installer
gh release upload "v$VERSION" "$TARGET_MSI" --clobber
echo -e "${GREEN}✅ Uploaded MSI installer${NC}"

# Upload portable ZIP
gh release upload "v$VERSION" "$PORTABLE_ZIP" --clobber
echo -e "${GREEN}✅ Uploaded portable package${NC}"

# Upload version info
gh release upload "v$VERSION" "$VERSION_FILE" --clobber
echo -e "${GREEN}✅ Uploaded version information${NC}"

echo -e "\n${GREEN}🎉 DEPLOYMENT SUCCESSFUL!${NC}"
echo -e "${CYAN}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
echo -e "${NC}📦 Release: v$VERSION${NC}"
echo -e "${NC}🔗 URL: https://github.com/ItMeDiaTech/Bulk_Editor_WPF/releases/tag/v$VERSION${NC}"
echo -e "${NC}📂 Assets:${NC}"
echo -e "${BLUE}   • BulkEditor-Setup-$VERSION.msi${NC}"
echo -e "${BLUE}   • BulkEditor-v$VERSION-Portable.zip${NC}"
echo -e "${BLUE}   • version.json${NC}"
echo -e "${CYAN}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"

# Show file sizes
if command -v du >/dev/null 2>&1; then
    MSI_SIZE=$(du -m "$TARGET_MSI" | cut -f1)
    ZIP_SIZE=$(du -m "$PORTABLE_ZIP" | cut -f1)
    echo -e "${NC}📊 Package Sizes:${NC}"
    echo -e "${BLUE}   • MSI: ${MSI_SIZE} MB${NC}"
    echo -e "${BLUE}   • ZIP: ${ZIP_SIZE} MB${NC}"
fi

echo -e "\n${GREEN}🎯 Deployment completed successfully! Ready for distribution.${NC}"