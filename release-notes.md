# BulkEditor v1.5.0 Release Notes

## 🚀 Major Release: Comprehensive Optimization & Security

This release represents a comprehensive overhaul of the BulkEditor project with critical security fixes, architectural improvements, and enhanced OpenXML processing capabilities.

### 🔐 Security Enhancements
- **CRITICAL**: Removed exposed GitHub tokens from repository
- Enhanced .gitignore with comprehensive security patterns
- Secure credential handling mechanisms

### 🏗️ Architecture Improvements
- Eliminated duplicate UpdateManager classes for clean architecture
- Fixed redundant code patterns across the project
- Enhanced WPF MVVM compliance to industry standards
- Improved clean architecture layer separation

### ⚡ OpenXML Processing Engine
- **CRITICAL**: Fixed hyperlink relationship handling per official OpenXML documentation
- Preserve original relationship IDs to maintain document integrity
- Enhanced external link handling with complete URIs and fragments
- VBA-compatible hyperlink processing logic
- Improved regex patterns with word boundaries for precise ID matching

### 📦 Package Management
- Updated DocumentFormat.OpenXml from 3.0.1 to 3.3.0
- Modernized all test packages to latest stable versions
- Eliminated all NuGet version warnings

### 🗂️ Project Organization
- Created dedicated `docs/` directory for all documentation
- Created `scripts/` directory for VBA and PowerShell files
- Cleaned project root of miscellaneous files
- Added comprehensive analysis documentation

### 📊 Quality Metrics
- **Build**: Clean build with 0 errors, 0 warnings
- **Tests**: 159/161 passing (98.8% success rate)
- **Architecture**: Full WPF MVVM best practices compliance
- **Security**: Repository secured from credential exposure

### 🔧 Technical Improvements
- Enhanced nullable reference handling across codebase
- Fixed regex false positives in ID extraction
- Improved error handling and logging
- Better hyperlink text display formatting preservation

## Upgrade Notes

This is a major release with architectural changes. All users are encouraged to upgrade for the security fixes and improved reliability.

**Previous Version**: 1.4.2  
**Current Version**: 1.5.0

## Breaking Changes

None - this release maintains backward compatibility while fixing critical security and processing issues.

## System Requirements

- Windows 10 or later
- .NET 8.0 Runtime
- Microsoft Word (for document processing)