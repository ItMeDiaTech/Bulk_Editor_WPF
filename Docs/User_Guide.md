# BulkEditor User Guide

## Table of Contents

1. [Getting Started](#getting-started)
2. [Main Interface](#main-interface)
3. [Document Processing](#document-processing)
4. [Replacement Features](#replacement-features)
5. [Settings Configuration](#settings-configuration)
6. [Troubleshooting](#troubleshooting)

## Getting Started

### System Requirements

- **Operating System**: Windows 10/11 (64-bit)
- **Framework**: .NET 8.0 Runtime (included with application)
- **Microsoft Word**: Required for document processing
- **Memory**: Minimum 4GB RAM (8GB recommended for large batch operations)
- **Storage**: 200MB free space for application and logs

### Installation

1. **Download** the BulkEditor package from the distribution location
2. **Extract** all files to your preferred installation directory
3. **Run** `BulkEditor.UI.exe` to start the application
4. **Configure** initial settings via the Settings menu

### First Launch

When you first start BulkEditor, the application will:

- Create default configuration files
- Initialize logging in the `Logs` directory
- Create a `Backups` directory for document safety
- Display the main processing interface

## Main Interface

### File Selection

**Drag & Drop**: Simply drag Word documents (.docx, .docm) from Windows Explorer onto the application window.

**Browse Button**: Click "Browse Documents" to select files using the standard Windows file dialog.

**File List**: Selected documents appear in the main list with status indicators:

- ⏳ **Pending**: Document queued for processing
- ✅ **Completed**: Successfully processed
- ❌ **Failed**: Processing encountered errors
- 🔄 **Processing**: Currently being processed

### Processing Controls

**Process Documents**: Starts batch processing of all selected documents.

**Clear List**: Removes all documents from the processing queue.

**Export Results**: Saves processing results to JSON or CSV format.

### Status Bar

The status bar provides real-time feedback:

- **Ready**: Application ready for document selection
- **Processing**: Shows current operation and progress percentage
- **Success**: Processing completed successfully
- **Warning**: Completed with warnings or issues
- **Error**: Processing failed with error details

## Document Processing

### Processing Pipeline

BulkEditor processes documents through several stages:

1. **Backup Creation**: Automatic backup before any modifications
2. **Metadata Extraction**: Document properties, word count, author info
3. **Hyperlink Extraction**: All hyperlinks with display text and URLs
4. **Hyperlink Validation**: URL accessibility and expiration checking
5. **Title Comparison**: Compare current titles with API responses
6. **Replacement Processing**: Apply configured replacement rules
7. **Text Optimization**: Clean up formatting and whitespace
8. **Change Logging**: Detailed audit trail of all modifications

### Hyperlink Management

**Lookup ID Detection**: Automatically detects TSRC and CMS lookup IDs using pattern:

```
(TSRC-[^-]+-[0-9]{6}|CMS-[^-]+-[0-9]{6})
```

**Content ID Generation**: Creates 6-digit Content IDs from lookup IDs for URL building.

**URL Updates**: Updates hyperlink URLs using pattern:

```
https://example.com/content/{ContentId}
```

**Title Validation**: Compares hyperlink titles with API responses and handles differences based on settings.

## Replacement Features

### Hyperlink Replacement

**Purpose**: Replace hyperlink display text based on exact title matches.

**Configuration**:

1. Navigate to Settings → Replacement
2. Enable "Enable hyperlink replacement"
3. Add replacement rules with:
   - **Title to Match**: Exact title text (case-insensitive)
   - **Content ID**: 6-digit Content ID for the replacement

**Behavior**: When a hyperlink's title (excluding Content ID) matches a rule, it's replaced with the format: `"API Title (Content_ID)"`

### Text Replacement

**Purpose**: Replace text throughout document content with intelligent capitalization preservation.

**Configuration**:

1. Navigate to Settings → Replacement
2. Enable "Enable text replacement"
3. Add replacement rules with:
   - **Source Text**: Text to find and replace
   - **Replacement Text**: New text to insert

**Capitalization Rules**:

- **ALL UPPERCASE** → Replacement becomes **ALL UPPERCASE**
- **all lowercase** → Replacement becomes **all lowercase**
- **First Letter Capitalized** → **First Letter Capitalized** in replacement
- **Mixed Case** → Preserves user's original capitalization

### Automatic Title Replacement

**Purpose**: Automatically update hyperlink titles based on API responses (like Base_File.vba).

**Configuration**:

1. Navigate to Settings → Validation
2. Choose behavior:
   - **Auto-replace titles with API response**: Automatically updates titles
   - **Report title differences in changelog**: Only reports differences

**Behavior**:

- Compares current hyperlink title (without Content ID) with API response
- If different and auto-replace enabled: Updates to API title + Content ID
- If different and reporting enabled: Logs "Possible Title Change" in changelog

## Settings Configuration

### Processing Settings

**Concurrency Control**:

- **Maximum Concurrent Documents**: Limits parallel processing (1-100)
- **Batch Size**: Documents processed per batch (10-500)

**Processing Options**:

- **Create backup before processing**: Safety backup creation
- **Validate hyperlinks**: Enable HTTP accessibility checks
- **Update hyperlinks automatically**: Apply URL updates
- **Add Content IDs**: Append Content IDs to display text
- **Optimize text formatting**: Clean up document formatting

### Validation Settings

**HTTP Configuration**:

- **HTTP Timeout**: Maximum wait time for HTTP requests (1-300 seconds)
- **Max Retry Attempts**: Number of retry attempts for failed requests (0-10)
- **User Agent**: HTTP User-Agent string for requests

**Title Handling**:

- **Auto-replace titles with API response**: Automatically update mismatched titles
- **Report title differences in changelog**: Log title differences without changing

### Replacement Settings

**Rule Management**:

- **Add Rule**: Create new replacement rules
- **Remove**: Delete selected rules
- **Clear Invalid**: Remove rules with blank fields
- **Enable/Disable**: Toggle individual rules

**Validation**:

- Rules with blank fields are automatically cleared
- Content IDs must contain 6-digit numbers
- Text rules cannot have identical source and replacement text

### Backup Settings

**Directory Configuration**:

- **Backup Directory**: Location for backup files
- **Create timestamped backups**: Include timestamp in backup names

**Cleanup Options**:

- **Auto cleanup old backups**: Remove backups older than specified age
- **Max backup age**: Retention period in days (1-365)

## Troubleshooting

### Common Issues

**"Document not found" Error**:

- Verify file path is correct
- Ensure file hasn't been moved or deleted
- Check file permissions

**"File is not a valid Word document" Error**:

- Ensure file has .docx or .docm extension
- Verify file isn't corrupted
- Try opening in Microsoft Word first

**"Processing failed" Error**:

- Check document isn't open in Word
- Verify sufficient disk space for backups
- Review detailed error logs in Logs directory

**Hyperlink Validation Failures**:

- Check internet connectivity
- Verify HTTP timeout settings aren't too low
- Review skip domains list for excluded URLs

### Performance Optimization

**For Large Batches**:

- Reduce maximum concurrent documents (try 5-10)
- Increase HTTP timeout for slow networks
- Monitor system memory usage

**For Slow Networks**:

- Increase HTTP timeout and retry attempts
- Consider processing smaller batches
- Enable caching for repeated lookups

### Log Files

Application logs are stored in the `Logs` directory with detailed information:

- **Information**: Normal processing events
- **Warning**: Non-critical issues that don't stop processing
- **Error**: Critical errors requiring attention
- **Debug**: Detailed diagnostic information (when debug logging enabled)

### Backup Recovery

If processing causes issues:

1. Navigate to the document's directory
2. Look for the `Backups` folder
3. Find the timestamped backup of your document
4. Restore by copying the backup over the processed file

### Settings Reset

To reset all settings to defaults:

1. Open Settings window
2. Click "Reset to Defaults" button
3. Click "Save Settings" to apply

Alternatively, delete the `appsettings.json` file and restart the application.

### Cache Management

Cache is automatically managed, but you can:

- Monitor cache performance via API statistics
- Cache automatically expires after 30 minutes
- Application restart clears all cached data

## Advanced Usage

### Command Line Integration

While BulkEditor is primarily a GUI application, you can integrate it with scripts:

```powershell
# Launch with specific documents
Start-Process "BulkEditor.UI.exe" -ArgumentList "document1.docx document2.docx"
```

### Batch Automation

For automated workflows:

1. Configure all settings via the UI
2. Use the API interfaces for programmatic access
3. Monitor processing via the progress reporting system

### Integration with Document Workflows

BulkEditor integrates well with:

- **SharePoint**: Process documents from SharePoint libraries
- **Network Drives**: Batch process shared documents
- **Version Control**: Process documents as part of content updates
- **Content Management**: Automate hyperlink maintenance workflows

## Support

For technical support:

- Review log files in the `Logs` directory for detailed error information
- Check the `Project_Info.md` file for development and architecture details
- Refer to the `API_Documentation.md` for programmatic integration
