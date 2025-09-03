# WPF Bulk Editor - Technical Specifications

## 🎯 **Business Logic Analysis (from Base_File.vba)**

### **Core Processing Workflow**

1. **Document Validation**: Remove invisible external hyperlinks
2. **ID Extraction**: Extract unique Lookup_IDs using regex pattern `(TSRC-[^-]+-[0-9]{6}|CMS-[^-]+-[0-9]{6})`
3. **API Communication**: POST JSON request with Lookup_ID array
4. **Response Processing**: Parse JSON response for document metadata
5. **Hyperlink Updates**: Update URLs, append Content IDs, mark expired/not found
6. **Changelog Generation**: Create structured log files with categorized results

### **Content ID Management**

- **Format**: Last 6 digits of Content_ID appended as " (123456)"
- **Logic**: Replace existing 5-digit format with 6-digit format if detected
- **Validation**: Only append if not already present

### **URL Structure**

- **Base URL**: `https://thesource.cvshealth.com/nuxeo/thesource/`
- **SubAddress**: `!/view?docid={Document_ID}`
- **Target Assembly**: Full URL = Base + "#" + SubAddress

### **Status Indicators**

- **Expired Documents**: Append " - Expired" to TextToDisplay
- **Not Found Documents**: Append " - Not Found" to TextToDisplay
- **Title Changes**: Detect and log when title differs from API response

### **Changelog Categories**

1. **Updated Links**: URL changes and Content ID additions
2. **Found Expired**: Documents marked as expired
3. **Not Found**: Documents not found in API response
4. **Found Error**: Invisible hyperlinks removed
5. **Potential Outdated Titles**: Title mismatches detected

## 🏗️ **Architecture Specifications**

### **Project Structure (Clean Architecture)**

```
BulkEditor/
├── BulkEditor.Core/                    # Domain Layer
│   ├── Models/                         # Domain Models
│   │   ├── Document.cs
│   │   ├── Hyperlink.cs
│   │   ├── ProcessingResult.cs
│   │   ├── ApiResponse.cs
│   │   └── ChangelogEntry.cs
│   ├── Enums/                          # Domain Enums
│   │   ├── DocumentStatus.cs
│   │   ├── HyperlinkStatus.cs
│   │   └── ProcessingOperation.cs
│   └── Exceptions/                     # Domain Exceptions
│       ├── DocumentProcessingException.cs
│       └── ApiCommunicationException.cs
├── BulkEditor.Application/             # Application Layer
│   ├── Services/                       # Application Services
│   │   ├── Abstractions/              # Service Interfaces
│   │   │   ├── IDocumentProcessingService.cs
│   │   │   ├── IApiService.cs
│   │   │   ├── IHyperlinkService.cs
│   │   │   ├── IChangelogService.cs
│   │   │   ├── IFileSystemService.cs
│   │   │   ├── IConfigurationService.cs
│   │   │   └── IProgressReportingService.cs
│   │   └── Implementations/           # Service Implementations
│   │       ├── DocumentProcessingService.cs
│   │       ├── ApiService.cs
│   │       ├── HyperlinkService.cs
│   │       ├── ChangelogService.cs
│   │       ├── FileSystemService.cs
│   │       ├── ConfigurationService.cs
│   │       └── ProgressReportingService.cs
│   ├── DTOs/                          # Data Transfer Objects
│   │   ├── ProcessingRequest.cs
│   │   ├── ApiRequest.cs
│   │   ├── ApiResponseDto.cs
│   │   └── ProgressUpdate.cs
│   └── Extensions/                    # Application Extensions
│       ├── DocumentExtensions.cs
│       └── HyperlinkExtensions.cs
├── BulkEditor.Infrastructure/          # Infrastructure Layer
│   ├── Configuration/                 # Configuration Models
│   │   ├── AppSettings.cs
│   │   ├── UserSettings.cs
│   │   ├── ApiSettings.cs
│   │   └── LoggingSettings.cs
│   ├── Logging/                       # Logging Implementation
│   │   ├── LoggingConfiguration.cs
│   │   └── StructuredLogger.cs
│   ├── Http/                          # HTTP Client Configuration
│   │   ├── ApiClient.cs
│   │   └── HttpClientConfiguration.cs
│   └── FileSystem/                    # File System Operations
│       ├── BackupManager.cs
│       ├── DocumentReader.cs
│       └── DocumentWriter.cs
├── BulkEditor.UI/                     # Presentation Layer (WPF)
│   ├── Views/                         # XAML Views
│   │   ├── MainWindow.xaml
│   │   ├── SettingsWindow.xaml
│   │   ├── ProgressWindow.xaml
│   │   └── LogViewWindow.xaml
│   ├── ViewModels/                    # View Models
│   │   ├── MainWindowViewModel.cs
│   │   ├── SettingsWindowViewModel.cs
│   │   ├── ProgressWindowViewModel.cs
│   │   └── LogViewWindowViewModel.cs
│   ├── Controls/                      # Custom Controls
│   │   ├── FileDropZone.xaml
│   │   ├── ProgressIndicator.xaml
│   │   └── ThemeToggle.xaml
│   ├── Converters/                    # Value Converters
│   │   ├── StatusToColorConverter.cs
│   │   ├── BoolToVisibilityConverter.cs
│   │   └── ProgressToPercentageConverter.cs
│   ├── Resources/                     # UI Resources
│   │   ├── Styles/
│   │   │   ├── MaterialDesignStyles.xaml
│   │   │   └── CustomStyles.xaml
│   │   ├── Themes/
│   │   │   ├── LightTheme.xaml
│   │   │   └── DarkTheme.xaml
│   │   └── Icons/
│   └── App.xaml                       # Application Entry Point
└── BulkEditor.Tests/                  # Test Layer
    ├── Unit/                          # Unit Tests
    ├── Integration/                   # Integration Tests
    └── TestData/                      # Test Documents
```

### **Service Layer Design**

#### **IDocumentProcessingService**

```csharp
public interface IDocumentProcessingService
{
    Task<ProcessingResult> ProcessDocumentAsync(string filePath, CancellationToken cancellationToken = default);
    Task<ProcessingResult> ProcessDocumentsAsync(IEnumerable<string> filePaths, IProgress<ProgressUpdate> progress = null, CancellationToken cancellationToken = default);
    Task<bool> ValidateDocumentAsync(string filePath);
    Task<BackupInfo> CreateBackupAsync(string filePath);
}
```

#### **IApiService**

```csharp
public interface IApiService
{
    Task<ApiResponseDto> GetDocumentMetadataAsync(IEnumerable<string> lookupIds, CancellationToken cancellationToken = default);
    Task<bool> ValidateApiConnectionAsync();
    Task<VersionInfo> GetVersionInfoAsync();
}
```

#### **IHyperlinkService**

```csharp
public interface IHyperlinkService
{
    IEnumerable<string> ExtractLookupIds(DocumentFormat.OpenXml.Wordprocessing.Document document);
    Task UpdateHyperlinksAsync(DocumentFormat.OpenXml.Wordprocessing.Document document, ApiResponseDto apiResponse);
    IEnumerable<HyperlinkInfo> GetInvisibleHyperlinks(DocumentFormat.OpenXml.Wordprocessing.Document document);
    void RemoveInvisibleHyperlinks(DocumentFormat.OpenXml.Wordprocessing.Document document);
}
```

### **Configuration Management**

#### **AppSettings.json**

```json
{
  "ApiSettings": {
    "BaseUrl": "https://api.example.com",
    "Timeout": 30,
    "RetryCount": 3,
    "RetryDelay": 1000
  },
  "ProcessingSettings": {
    "MaxConcurrentDocuments": 10,
    "BackupEnabled": true,
    "BackupRetentionDays": 30
  },
  "LoggingSettings": {
    "MinimumLevel": "Information",
    "FileSizeLimit": 50,
    "RetainedFileCountLimit": 10
  }
}
```

#### **UserSettings.json** (AppData)

```json
{
  "ThemeSettings": {
    "CurrentTheme": "Light",
    "AutoDetectSystemTheme": true
  },
  "UISettings": {
    "WindowWidth": 1200,
    "WindowHeight": 800,
    "WindowState": "Normal"
  },
  "ProcessingSettings": {
    "DefaultInputFolder": "",
    "DefaultOutputFolder": "",
    "AutoOpenChangelog": true
  }
}
```

### **MVVM Implementation with CommunityToolkit.Mvvm**

#### **Base ViewModel**

```csharp
public partial class BaseViewModel : ObservableObject
{
    [ObservableProperty]
    private bool isBusy;

    [ObservableProperty]
    private string statusMessage = string.Empty;

    protected readonly ILogger logger;

    public BaseViewModel(ILogger logger)
    {
        this.logger = logger;
    }
}
```

#### **MainWindowViewModel**

```csharp
public partial class MainWindowViewModel : BaseViewModel
{
    private readonly IDocumentProcessingService processingService;
    private readonly IConfigurationService configurationService;

    [ObservableProperty]
    private ObservableCollection<DocumentItem> documents = new();

    [ObservableProperty]
    private bool isProcessing;

    [ObservableProperty]
    private double progressPercentage;

    [RelayCommand]
    private async Task SelectFilesAsync()
    {
        // File selection logic
    }

    [RelayCommand]
    private async Task ProcessDocumentsAsync()
    {
        // Document processing logic
    }

    [RelayCommand]
    private async Task OpenSettingsAsync()
    {
        // Settings window logic
    }
}
```

### **Error Handling Strategy**

#### **Global Exception Handling**

```csharp
public class GlobalExceptionHandler
{
    private readonly ILogger logger;

    public void HandleException(Exception exception, string context)
    {
        logger.LogError(exception, "Unhandled exception in {Context}", context);

        // Show user-friendly error message
        // Log technical details
        // Attempt recovery if possible
    }
}
```

#### **Structured Error Response**

```csharp
public class ProcessingResult
{
    public bool IsSuccess { get; set; }
    public string Message { get; set; }
    public List<string> Errors { get; set; } = new();
    public List<string> Warnings { get; set; } = new();
    public ProcessingStatistics Statistics { get; set; }
}
```

### **Progress Reporting Implementation**

#### **Progress Update Model**

```csharp
public class ProgressUpdate
{
    public int CurrentDocument { get; set; }
    public int TotalDocuments { get; set; }
    public string CurrentOperation { get; set; }
    public string DocumentName { get; set; }
    public double PercentageComplete => (double)CurrentDocument / TotalDocuments * 100;
}
```

#### **IProgress<T> Implementation**

```csharp
public class ProgressReportingService : IProgressReportingService
{
    public IProgress<ProgressUpdate> CreateProgress(Action<ProgressUpdate> callback)
    {
        return new Progress<ProgressUpdate>(callback);
    }
}
```

### **Logging Implementation with Serilog**

#### **Structured Logging Configuration**

```csharp
public static class LoggingConfiguration
{
    public static ILogger CreateLogger(LoggingSettings settings)
    {
        return new LoggerConfiguration()
            .MinimumLevel.Is(settings.MinimumLevel)
            .Enrich.WithProperty("Application", "BulkEditor")
            .Enrich.WithMachineName()
            .Enrich.WithUserName()
            .WriteTo.Console(outputTemplate: "[{Timestamp:HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}")
            .WriteTo.File(
                path: Path.Combine(settings.LogDirectory, "log-.txt"),
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: settings.RetainedFileCountLimit,
                fileSizeLimitBytes: settings.FileSizeLimit * 1024 * 1024,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
            .CreateLogger();
    }
}
```

### **Material Design UI Implementation**

#### **Theme Management**

```csharp
public class ThemeService : IThemeService
{
    public void ApplyTheme(ThemeType themeType)
    {
        var theme = themeType == ThemeType.Dark ? BaseTheme.Dark : BaseTheme.Light;

        var helper = new PaletteHelper();
        helper.SetTheme(Theme.Create(theme, Colors.Blue, Colors.Orange));
    }

    public void DetectSystemTheme()
    {
        // Windows 10/11 theme detection logic
    }
}
```

### **Deployment Strategy**

#### **AppData Installation Structure**

```
%APPDATA%/BulkEditor/
├── BulkEditor.exe
├── appsettings.json
├── UserSettings.json
├── Logs/
├── Backups/
├── Temp/
└── Dependencies/
```

#### **Installer Requirements**

- **Target**: Windows 10/11 x64
- **Framework**: .NET 8 Runtime (self-contained deployment)
- **Permissions**: User-level installation (no admin required)
- **Uninstaller**: Standard Windows uninstall support

## 🔄 **Implementation Phases**

### **Phase 1: Foundation (Week 1)**

1. Project structure setup
2. Dependency injection configuration
3. Basic MVVM infrastructure
4. Configuration management
5. Logging implementation

### **Phase 2: Core Services (Week 2)**

1. Document processing service
2. API communication service
3. Hyperlink management service
4. File system operations
5. Unit tests for core services

### **Phase 3: UI Implementation (Week 3)**

1. Main window with Material Design
2. Progress reporting UI
3. Settings management
4. Theme switching
5. Drag & drop functionality

### **Phase 4: Integration & Testing (Week 4)**

1. End-to-end integration
2. Batch processing workflow
3. Error handling validation
4. Performance optimization
5. Deployment preparation

## ✅ **Success Criteria Validation**

### **Performance Requirements**

- **Concurrent Documents**: Support 200 documents simultaneously
- **Memory Usage**: < 1GB for typical batch operations
- **Response Time**: < 30 seconds for API calls with retry logic
- **UI Responsiveness**: < 100ms UI freeze during operations

### **Reliability Requirements**

- **Error Recovery**: Graceful handling of document corruption, network failures
- **Data Integrity**: Atomic file operations with rollback capability
- **Audit Trail**: Comprehensive logging of all operations
- **Backup Strategy**: Automatic backup creation with configurable retention

### **User Experience Requirements**

- **Intuitive Interface**: Clear workflow with minimal clicks
- **Progress Feedback**: Real-time progress with cancellation support
- **Accessibility**: Keyboard navigation, screen reader support
- **Theme Support**: System theme integration with manual override

This technical specification provides the blueprint for implementing a production-ready WPF Bulk Editor that maintains the core business logic from the VBA implementation while leveraging modern .NET architecture patterns.
