# BulkEditor API Documentation

## Overview

The BulkEditor application provides a comprehensive API for document processing, hyperlink management, and text replacement operations. This documentation covers the core interfaces and services available for integration and extension.

## Core Interfaces

### IDocumentProcessor

The main interface for document processing operations.

#### Methods

```csharp
Task<Document> ProcessDocumentAsync(string filePath, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
```

Processes a single document with optional progress reporting and cancellation support.

**Parameters:**

- `filePath`: Path to the Word document (.docx/.docm)
- `progress`: Optional progress reporting callback
- `cancellationToken`: Cancellation token for async operations

**Returns:** Processed document with metadata, hyperlinks, and change log

```csharp
Task<IEnumerable<Document>> ProcessDocumentsBatchAsync(IEnumerable<string> filePaths, IProgress<BatchProcessingProgress>? progress = null, CancellationToken cancellationToken = default)
```

Processes multiple documents concurrently with batch progress tracking.

### IHyperlinkValidator

Interface for hyperlink validation and Content ID management.

#### Methods

```csharp
Task<HyperlinkValidationResult> ValidateHyperlinkAsync(Hyperlink hyperlink, CancellationToken cancellationToken = default)
```

Validates a single hyperlink and performs title comparison with API response.

```csharp
string ExtractLookupId(string url)
```

Extracts lookup ID from URL using regex pattern `(TSRC-[^-]+-[0-9]{6}|CMS-[^-]+-[0-9]{6})`.

```csharp
Task<string> GenerateContentIdAsync(string lookupId, CancellationToken cancellationToken = default)
```

Generates 6-digit Content ID from lookup ID with caching support.

### IReplacementService

Interface for hyperlink and text replacement operations.

#### Methods

```csharp
Task<Document> ProcessReplacementsAsync(Document document, CancellationToken cancellationToken = default)
```

Processes all enabled replacement rules for a document.

```csharp
Task<ReplacementValidationResult> ValidateReplacementRulesAsync(IEnumerable<object> rules, CancellationToken cancellationToken = default)
```

Validates replacement rules for correctness and format compliance.

### IHyperlinkReplacementService

Specialized interface for hyperlink replacement with title matching.

#### Methods

```csharp
Task<Document> ProcessHyperlinkReplacementsAsync(Document document, IEnumerable<HyperlinkReplacementRule> rules, CancellationToken cancellationToken = default)
```

Processes hyperlink replacements using case-insensitive title matching.

```csharp
Task<string> LookupTitleByContentIdAsync(string contentId, CancellationToken cancellationToken = default)
```

Looks up document title by Content ID (simulated API call).

```csharp
string BuildUrlFromContentId(string contentId)
```

Builds URL using pattern: `https://example.com/content/{contentId}`

### ITextReplacementService

Interface for text replacement with capitalization preservation.

#### Methods

```csharp
Task<Document> ProcessTextReplacementsAsync(Document document, IEnumerable<TextReplacementRule> rules, CancellationToken cancellationToken = default)
```

Processes text replacements throughout document content.

```csharp
string ReplaceTextWithCapitalizationPreservation(string sourceText, string searchText, string replacementText)
```

Performs intelligent text replacement preserving original capitalization patterns.

### ICacheService

Performance optimization interface for caching API responses and computed values.

#### Methods

```csharp
Task<T> GetOrSetAsync<T>(string key, Func<Task<T>> factory, TimeSpan? expiry = null, CancellationToken cancellationToken = default)
```

Gets cached value or computes and stores new value.

```csharp
CacheStatistics GetStatistics()
```

Returns cache performance metrics including hit ratio and memory usage.

## Configuration Models

### AppSettings

Main configuration object with the following sections:

#### ProcessingSettings

- `MaxConcurrentDocuments`: Maximum concurrent document processing (default: 200)
- `BatchSize`: Documents per batch (default: 50)
- `ValidateHyperlinks`: Enable hyperlink validation (default: true)
- `UpdateHyperlinks`: Enable automatic hyperlink updates (default: true)
- `OptimizeText`: Enable text optimization (default: false)
- `LookupIdPattern`: Regex pattern for lookup ID extraction

#### ValidationSettings

- `HttpTimeout`: HTTP request timeout (default: 30 seconds)
- `MaxRetryAttempts`: Maximum retry attempts (default: 3)
- `CheckExpiredContent`: Check for expired content indicators (default: true)
- `AutoReplaceTitles`: Automatically replace titles with API response (default: false)
- `ReportTitleDifferences`: Report title differences in changelog (default: true)

#### ReplacementSettings

- `EnableHyperlinkReplacement`: Enable hyperlink replacement rules (default: false)
- `EnableTextReplacement`: Enable text replacement rules (default: false)
- `HyperlinkRules`: Collection of hyperlink replacement rules
- `TextRules`: Collection of text replacement rules
- `MaxReplacementRules`: Maximum number of rules (default: 50)

#### BackupSettings

- `BackupDirectory`: Directory for backup files (default: "Backups")
- `CreateTimestampedBackups`: Include timestamp in backup filenames (default: true)
- `MaxBackupAge`: Maximum backup retention in days (default: 30)

## Entity Models

### Document

Represents a processed Word document with metadata and change tracking.

**Properties:**

- `FilePath`: Full path to document file
- `FileName`: Document filename
- `Status`: Processing status (Processing, Completed, Failed)
- `Metadata`: Document metadata (title, author, word count, etc.)
- `Hyperlinks`: Collection of extracted hyperlinks
- `ChangeLog`: Detailed change tracking information
- `ProcessingErrors`: Collection of processing errors
- `BackupPath`: Path to backup file

### Hyperlink

Represents a hyperlink within a document.

**Properties:**

- `Id`: Unique hyperlink identifier
- `OriginalUrl`: Original hyperlink URL
- `UpdatedUrl`: Updated URL after processing
- `DisplayText`: Hyperlink display text
- `LookupId`: Extracted lookup ID (TSRC/CMS pattern)
- `ContentId`: Generated 6-digit Content ID
- `Status`: Validation status (Valid, Invalid, Expired, NotFound)
- `RequiresUpdate`: Indicates if hyperlink needs updating
- `ActionTaken`: Action performed on hyperlink

### TitleComparisonResult

Result of comparing hyperlink title with API response.

**Properties:**

- `TitlesDiffer`: Whether titles differ between current and API
- `CurrentTitle`: Current hyperlink title (without Content ID)
- `ApiTitle`: Title from API response
- `ContentId`: Associated Content ID
- `WasReplaced`: Whether title was automatically replaced
- `ActionTaken`: Description of action performed

## Usage Examples

### Basic Document Processing

```csharp
// Inject document processor via DI
var processor = serviceProvider.GetRequiredService<IDocumentProcessor>();

// Process single document
var document = await processor.ProcessDocumentAsync("path/to/document.docx");

// Check results
if (document.Status == DocumentStatus.Completed)
{
    Console.WriteLine($"Processed {document.Hyperlinks.Count} hyperlinks");
    Console.WriteLine($"Changes: {document.ChangeLog.Summary}");
}
```

### Replacement Rules Configuration

```csharp
// Configure hyperlink replacement
var hyperlinkRule = new HyperlinkReplacementRule
{
    TitleToMatch = "Old Document Title",
    ContentId = "123456",
    IsEnabled = true
};

// Configure text replacement
var textRule = new TextReplacementRule
{
    SourceText = "legacy text",
    ReplacementText = "modern content",
    IsEnabled = true
};

// Add to settings
appSettings.Replacement.HyperlinkRules.Add(hyperlinkRule);
appSettings.Replacement.TextRules.Add(textRule);
appSettings.Replacement.EnableHyperlinkReplacement = true;
appSettings.Replacement.EnableTextReplacement = true;
```

### Batch Processing with Progress

```csharp
var filePaths = new[] { "doc1.docx", "doc2.docx", "doc3.docx" };

var progressReporter = new Progress<BatchProcessingProgress>(progress =>
{
    Console.WriteLine($"Processing: {progress.CurrentDocument}");
    Console.WriteLine($"Progress: {progress.PercentageComplete:F1}%");
    Console.WriteLine($"Completed: {progress.ProcessedDocuments}/{progress.TotalDocuments}");
});

var results = await processor.ProcessDocumentsBatchAsync(filePaths, progressReporter);
```

### Cache Integration

```csharp
// Use cache service for performance optimization
var cacheService = serviceProvider.GetRequiredService<ICacheService>();

// Cache expensive operations
var title = await cacheService.GetOrSetAsync(
    $"title_lookup_{contentId}",
    async () => await apiService.GetTitleAsync(contentId),
    TimeSpan.FromHours(1)
);

// Check cache performance
var stats = cacheService.GetStatistics();
Console.WriteLine($"Cache hit ratio: {stats.HitRatio:P2}");
```

## Error Handling

All services implement comprehensive error handling with structured logging:

```csharp
try
{
    var document = await processor.ProcessDocumentAsync(filePath);
}
catch (FileNotFoundException ex)
{
    // Handle missing file
}
catch (InvalidOperationException ex)
{
    // Handle invalid document format
}
catch (Exception ex)
{
    // Handle general processing errors
    logger.LogError(ex, "Document processing failed");
}
```

## Extension Points

### Custom Replacement Services

Implement `IHyperlinkReplacementService` or `ITextReplacementService` for custom replacement logic:

```csharp
public class CustomHyperlinkReplacer : IHyperlinkReplacementService
{
    public async Task<Document> ProcessHyperlinkReplacementsAsync(
        Document document,
        IEnumerable<HyperlinkReplacementRule> rules,
        CancellationToken cancellationToken = default)
    {
        // Custom replacement implementation
        return document;
    }
}
```

### Custom Validation

Extend `IHyperlinkValidator` for custom validation logic:

```csharp
public class CustomValidator : IHyperlinkValidator
{
    public async Task<HyperlinkValidationResult> ValidateHyperlinkAsync(
        Hyperlink hyperlink,
        CancellationToken cancellationToken = default)
    {
        // Custom validation implementation
        return new HyperlinkValidationResult();
    }
}
```

## Performance Considerations

1. **Caching**: Content ID and title lookups are cached for 30 minutes by default
2. **Concurrency**: Batch processing uses configurable concurrency limits
3. **Memory Management**: Automatic garbage collection and cache cleanup
4. **Progress Reporting**: Non-blocking progress updates via `IProgress<T>`
5. **Cancellation**: Full `CancellationToken` support throughout the pipeline

## Security

- Input validation on all user-provided data
- Safe regex patterns with compiled options
- Backup creation before document modification
- Comprehensive error logging without sensitive data exposure
- Read-only document access during metadata extraction
