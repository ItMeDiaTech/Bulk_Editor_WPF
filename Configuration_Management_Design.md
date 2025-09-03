# WPF Bulk Editor - Configuration Management with Strongly-Typed Models

## 🎯 **Configuration Architecture Overview**

### **Configuration Layers**

- **Application Settings**: Core application configuration (appsettings.json)
- **User Settings**: Personal preferences and state (UserSettings.json in AppData)
- **Environment Settings**: Environment-specific overrides (appsettings.{Environment}.json)
- **Runtime Settings**: Temporary settings that don't persist between sessions

### **Configuration Flow**

```mermaid
graph TB
    AppSettings[appsettings.json]
    EnvSettings[appsettings.{Environment}.json]
    UserSettings[UserSettings.json - AppData]
    RuntimeConfig[Runtime Configuration]

    AppSettings --> ConfigService[Configuration Service]
    EnvSettings --> ConfigService
    UserSettings --> ConfigService
    ConfigService --> RuntimeConfig

    ConfigService --> Validation[Configuration Validation]
    Validation --> App[Application Services]

    subgraph "Configuration Sources (Priority Order)"
        P1[1. Runtime Settings]
        P2[2. User Settings]
        P3[3. Environment Settings]
        P4[4. Application Settings]
    end
```

## 📋 **Strongly-Typed Configuration Models**

### **Core Application Settings**

```csharp
/// <summary>
/// Root application configuration model
/// </summary>
public class AppSettings
{
    [Required]
    public ApiSettings Api { get; set; } = new();

    [Required]
    public ProcessingSettings Processing { get; set; } = new();

    [Required]
    public LoggingSettings Logging { get; set; } = new();

    [Required]
    public SecuritySettings Security { get; set; } = new();

    [Required]
    public PerformanceSettings Performance { get; set; } = new();
}

/// <summary>
/// API communication configuration
/// </summary>
public class ApiSettings
{
    [Required(ErrorMessage = "API Base URL is required")]
    [Url(ErrorMessage = "API Base URL must be a valid URL")]
    public string BaseUrl { get; set; } = "https://api.example.com";

    [Range(5, 300, ErrorMessage = "Timeout must be between 5 and 300 seconds")]
    public int TimeoutSeconds { get; set; } = 30;

    [Range(1, 10, ErrorMessage = "Retry count must be between 1 and 10")]
    public int RetryCount { get; set; } = 3;

    [Range(100, 10000, ErrorMessage = "Retry delay must be between 100 and 10000 milliseconds")]
    public int RetryDelayMs { get; set; } = 1000;

    [Range(5, 60, ErrorMessage = "Circuit breaker timeout must be between 5 and 60 seconds")]
    public int CircuitBreakerTimeoutSeconds { get; set; } = 30;

    [Range(3, 20, ErrorMessage = "Circuit breaker failure threshold must be between 3 and 20")]
    public int CircuitBreakerFailureThreshold { get; set; } = 5;

    public string? ApiKey { get; set; }

    public Dictionary<string, string> Headers { get; set; } = new();
}

/// <summary>
/// Document processing configuration
/// </summary>
public class ProcessingSettings
{
    [Range(1, 50, ErrorMessage = "Max concurrent documents must be between 1 and 50")]
    public int MaxConcurrentDocuments { get; set; } = 5;

    [Range(1, 200, ErrorMessage = "Batch size must be between 1 and 200")]
    public int BatchSize { get; set; } = 20;

    public bool CreateBackups { get; set; } = true;

    [Range(1, 365, ErrorMessage = "Backup retention must be between 1 and 365 days")]
    public int BackupRetentionDays { get; set; } = 30;

    [Required(ErrorMessage = "Backup location is required when backups are enabled")]
    public string BackupLocation { get; set; } = "%APPDATA%\\BulkEditor\\Backups";

    public bool AutoCleanupBackups { get; set; } = true;

    [Range(1, 100, ErrorMessage = "Max backup files must be between 1 and 100")]
    public int MaxBackupFiles { get; set; } = 50;

    public bool ValidateDocumentsBeforeProcessing { get; set; } = true;

    public bool CreateChangelogOnCompletion { get; set; } = true;

    public ChangelogFormat DefaultChangelogFormat { get; set; } = ChangelogFormat.Text;
}

/// <summary>
/// Structured logging configuration
/// </summary>
public class LoggingSettings
{
    [Required]
    public LogLevel MinimumLevel { get; set; } = LogLevel.Information;

    [Required(ErrorMessage = "Log directory is required")]
    public string LogDirectory { get; set; } = "%APPDATA%\\BulkEditor\\Logs";

    [Range(1, 1000, ErrorMessage = "File size limit must be between 1 and 1000 MB")]
    public int FileSizeLimitMB { get; set; } = 50;

    [Range(1, 100, ErrorMessage = "Retained file count must be between 1 and 100")]
    public int RetainedFileCountLimit { get; set; } = 10;

    public bool EnableConsoleLogging { get; set; } = true;

    public bool EnableFileLogging { get; set; } = true;

    public bool EnableStructuredLogging { get; set; } = true;

    public string LogFilePattern { get; set; } = "log-{Date}.txt";

    public string OutputTemplate { get; set; } =
        "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}";

    public Dictionary<string, LogLevel> LogLevelOverrides { get; set; } = new();
}

/// <summary>
/// Security and validation configuration
/// </summary>
public class SecuritySettings
{
    public bool ValidateFilePaths { get; set; } = true;

    public bool AllowNetworkPaths { get; set; } = false;

    public bool RequireFileExtensionValidation { get; set; } = true;

    [Required]
    public List<string> AllowedFileExtensions { get; set; } = new() { ".docx" };

    [Range(1, 1000, ErrorMessage = "Max file size must be between 1 and 1000 MB")]
    public int MaxFileSizeMB { get; set; } = 100;

    public bool SanitizeUserInputs { get; set; } = true;

    public bool LogSecurityEvents { get; set; } = true;
}

/// <summary>
/// Performance optimization configuration
/// </summary>
public class PerformanceSettings
{
    [Range(100, 10000, ErrorMessage = "Memory cache size must be between 100 and 10000 MB")]
    public int MemoryCacheSizeMB { get; set; } = 500;

    [Range(1, 60, ErrorMessage = "Cache expiration must be between 1 and 60 minutes")]
    public int CacheExpirationMinutes { get; set; } = 15;

    public bool EnableMetricsCollection { get; set; } = true;

    public bool EnablePerformanceCounters { get; set; } = false;

    [Range(1, 10, ErrorMessage = "GC collection frequency must be between 1 and 10")]
    public int GarbageCollectionFrequency { get; set; } = 3;

    public bool OptimizeForLargeFiles { get; set; } = false;
}
```

### **User Settings Model**

```csharp
/// <summary>
/// User-specific configuration and preferences
/// </summary>
public class UserSettings
{
    [Required]
    public UISettings UI { get; set; } = new();

    [Required]
    public WorkflowSettings Workflow { get; set; } = new();

    [Required]
    public NotificationSettings Notifications { get; set; } = new();

    public Dictionary<string, object> CustomSettings { get; set; } = new();

    public DateTime LastModified { get; set; } = DateTime.UtcNow;

    public string Version { get; set; } = "1.0.0";
}

/// <summary>
/// User interface preferences
/// </summary>
public class UISettings
{
    public ThemeType SelectedTheme { get; set; } = ThemeType.System;

    public bool AutoDetectSystemTheme { get; set; } = true;

    public string AccentColor { get; set; } = "Blue";

    [Range(800, 3840, ErrorMessage = "Window width must be between 800 and 3840 pixels")]
    public int WindowWidth { get; set; } = 1200;

    [Range(600, 2160, ErrorMessage = "Window height must be between 600 and 2160 pixels")]
    public int WindowHeight { get; set; } = 800;

    public WindowState WindowState { get; set; } = WindowState.Normal;

    public bool RememberWindowPosition { get; set; } = true;

    public double WindowLeft { get; set; } = 100;

    public double WindowTop { get; set; } = 100;

    public bool ShowStatusBar { get; set; } = true;

    public bool ShowProgressDetails { get; set; } = true;

    public bool AutoRefreshDocumentList { get; set; } = true;

    [Range(1, 100, ErrorMessage = "Recent files count must be between 1 and 100")]
    public int RecentFilesCount { get; set; } = 10;

    public List<string> RecentFiles { get; set; } = new();

    public List<string> RecentFolders { get; set; } = new();
}

/// <summary>
/// Workflow and processing preferences
/// </summary>
public class WorkflowSettings
{
    public string DefaultInputFolder { get; set; } = string.Empty;

    public string DefaultOutputFolder { get; set; } = string.Empty;

    public bool AutoOpenChangelog { get; set; } = true;

    public bool AutoSaveSettings { get; set; } = true;

    public bool ConfirmBeforeProcessing { get; set; } = true;

    public bool ConfirmBeforeDeleting { get; set; } = true;

    public bool AutoSelectAllFiles { get; set; } = false;

    public ProcessingMode DefaultProcessingMode { get; set; } = ProcessingMode.Batch;

    public bool SaveProcessingResults { get; set; } = true;

    public string ProcessingResultsFolder { get; set; } = "%APPDATA%\\BulkEditor\\Results";

    public bool CreateDetailedLogs { get; set; } = true;

    public FileNamingConvention FileNamingConvention { get; set; } = FileNamingConvention.Original;
}

/// <summary>
/// Notification and feedback preferences
/// </summary>
public class NotificationSettings
{
    public bool ShowCompletionNotifications { get; set; } = true;

    public bool ShowErrorNotifications { get; set; } = true;

    public bool ShowWarningNotifications { get; set; } = true;

    public bool ShowProgressNotifications { get; set; } = false;

    public bool PlaySoundsOnCompletion { get; set; } = false;

    public bool ShowSystemTrayNotifications { get; set; } = true;

    [Range(1, 60, ErrorMessage = "Notification duration must be between 1 and 60 seconds")]
    public int NotificationDurationSeconds { get; set; } = 5;

    public NotificationPosition NotificationPosition { get; set; } = NotificationPosition.BottomRight;
}
```

### **Configuration Enums**

```csharp
public enum ThemeType
{
    Light,
    Dark,
    System
}

public enum ChangelogFormat
{
    Text,
    Html,
    Markdown,
    Json
}

public enum ProcessingMode
{
    Sequential,
    Batch,
    Parallel
}

public enum FileNamingConvention
{
    Original,
    Timestamp,
    Incremental,
    Custom
}

public enum NotificationPosition
{
    TopLeft,
    TopRight,
    BottomLeft,
    BottomRight,
    Center
}

public enum WindowState
{
    Normal,
    Minimized,
    Maximized
}
```

## 🔧 **Configuration Service Implementation**

### **IConfigurationService Interface**

```csharp
/// <summary>
/// Comprehensive configuration management service
/// </summary>
public interface IConfigurationService
{
    /// <summary>
    /// Gets strongly-typed configuration section
    /// </summary>
    T GetConfiguration<T>(string? sectionName = null) where T : class, new();

    /// <summary>
    /// Updates and persists configuration section
    /// </summary>
    Task<ConfigurationResult> UpdateConfigurationAsync<T>(string sectionName, T configuration) where T : class;

    /// <summary>
    /// Validates configuration integrity
    /// </summary>
    Task<ValidationResult> ValidateConfigurationAsync();

    /// <summary>
    /// Resets configuration to defaults
    /// </summary>
    Task<ConfigurationResult> ResetToDefaultsAsync(string? sectionName = null);

    /// <summary>
    /// Exports configuration to file
    /// </summary>
    Task<ConfigurationResult> ExportConfigurationAsync(string filePath, ConfigurationScope scope = ConfigurationScope.All);

    /// <summary>
    /// Imports configuration from file
    /// </summary>
    Task<ConfigurationResult> ImportConfigurationAsync(string filePath, bool mergeWithExisting = true);

    /// <summary>
    /// Gets configuration change notifications
    /// </summary>
    IObservable<ConfigurationChangeNotification> GetChangeNotifications();

    /// <summary>
    /// Reloads configuration from all sources
    /// </summary>
    Task ReloadConfigurationAsync();

    /// <summary>
    /// Gets configuration metadata and statistics
    /// </summary>
    ConfigurationMetadata GetConfigurationMetadata();
}

public enum ConfigurationScope
{
    Application,
    User,
    All
}

public class ConfigurationResult
{
    public bool IsSuccess { get; set; }
    public string Message { get; set; } = string.Empty;
    public List<string> Errors { get; set; } = new();
    public List<string> Warnings { get; set; } = new();
}

public class ConfigurationChangeNotification
{
    public string SectionName { get; set; } = string.Empty;
    public string PropertyName { get; set; } = string.Empty;
    public object? OldValue { get; set; }
    public object? NewValue { get; set; }
    public DateTime Timestamp { get; set; } = DateTime.UtcNow;
    public ConfigurationSource Source { get; set; }
}

public enum ConfigurationSource
{
    Application,
    Environment,
    User,
    Runtime
}

public class ConfigurationMetadata
{
    public Dictionary<string, DateTime> LastModified { get; set; } = new();
    public Dictionary<string, string> SourceFiles { get; set; } = new();
    public Dictionary<string, ValidationResult> ValidationResults { get; set; } = new();
    public Dictionary<string, int> AccessCounts { get; set; } = new();
}
```

### **Configuration Service Implementation**

```csharp
/// <summary>
/// Production configuration service implementation
/// </summary>
public class ConfigurationService : IConfigurationService, IDisposable
{
    private readonly IConfiguration _configuration;
    private readonly ILogger<ConfigurationService> _logger;
    private readonly IOptionsMonitor<AppSettings> _appSettingsMonitor;
    private readonly Subject<ConfigurationChangeNotification> _changeNotifications = new();
    private readonly string _userSettingsPath;
    private readonly FileSystemWatcher _fileWatcher;
    private readonly ConcurrentDictionary<string, object> _configurationCache = new();
    private readonly SemaphoreSlim _updateSemaphore = new(1, 1);

    public ConfigurationService(
        IConfiguration configuration,
        ILogger<ConfigurationService> logger,
        IOptionsMonitor<AppSettings> appSettingsMonitor)
    {
        _configuration = configuration;
        _logger = logger;
        _appSettingsMonitor = appSettingsMonitor;

        _userSettingsPath = GetUserSettingsPath();
        _fileWatcher = SetupFileWatcher();

        // Monitor application settings changes
        _appSettingsMonitor.OnChange(OnAppSettingsChanged);
    }

    public T GetConfiguration<T>(string? sectionName = null) where T : class, new()
    {
        var cacheKey = GenerateCacheKey<T>(sectionName);

        if (_configurationCache.TryGetValue(cacheKey, out var cached))
        {
            return (T)cached;
        }

        var config = LoadConfiguration<T>(sectionName);
        _configurationCache.TryAdd(cacheKey, config);

        return config;
    }

    public async Task<ConfigurationResult> UpdateConfigurationAsync<T>(string sectionName, T configuration) where T : class
    {
        await _updateSemaphore.WaitAsync();

        try
        {
            // Validate configuration before saving
            var validationResult = await ValidateConfigurationObjectAsync(configuration);
            if (!validationResult.IsValid)
            {
                return new ConfigurationResult
                {
                    IsSuccess = false,
                    Message = "Configuration validation failed",
                    Errors = validationResult.Errors
                };
            }

            // Determine if this is user or application configuration
            var isUserConfig = IsUserConfiguration<T>();
            var filePath = isUserConfig ? _userSettingsPath : GetAppSettingsPath();

            // Load existing configuration
            var existingConfig = LoadConfigurationFromFile(filePath);

            // Update the specific section
            existingConfig[sectionName] = JToken.FromObject(configuration);

            // Save to file
            await SaveConfigurationToFileAsync(filePath, existingConfig);

            // Update cache
            var cacheKey = GenerateCacheKey<T>(sectionName);
            _configurationCache.AddOrUpdate(cacheKey, configuration, (key, old) => configuration);

            // Notify change
            NotifyConfigurationChanged(sectionName, configuration, GetConfigurationSource<T>());

            _logger.LogInformation("Configuration updated for section {SectionName}", sectionName);

            return new ConfigurationResult { IsSuccess = true, Message = "Configuration updated successfully" };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to update configuration for section {SectionName}", sectionName);
            return new ConfigurationResult
            {
                IsSuccess = false,
                Message = "Failed to update configuration",
                Errors = new List<string> { ex.Message }
            };
        }
        finally
        {
            _updateSemaphore.Release();
        }
    }

    public async Task<ValidationResult> ValidateConfigurationAsync()
    {
        var results = new List<ValidationResult>();

        // Validate application settings
        var appSettings = GetConfiguration<AppSettings>();
        results.Add(await ValidateConfigurationObjectAsync(appSettings));

        // Validate user settings
        var userSettings = GetConfiguration<UserSettings>();
        results.Add(await ValidateConfigurationObjectAsync(userSettings));

        // Aggregate results
        var aggregateResult = new ValidationResult
        {
            IsValid = results.All(r => r.IsValid),
            Errors = results.SelectMany(r => r.Errors).ToList(),
            Warnings = results.SelectMany(r => r.Warnings).ToList()
        };

        return aggregateResult;
    }

    public async Task<ConfigurationResult> ResetToDefaultsAsync(string? sectionName = null)
    {
        try
        {
            if (sectionName == null)
            {
                // Reset all user settings to defaults
                var defaultUserSettings = new UserSettings();
                await UpdateConfigurationAsync("UserSettings", defaultUserSettings);
            }
            else
            {
                // Reset specific section
                var defaultValue = CreateDefaultConfiguration(sectionName);
                if (defaultValue != null)
                {
                    await UpdateConfigurationAsync(sectionName, defaultValue);
                }
            }

            _logger.LogInformation("Configuration reset to defaults for section {SectionName}", sectionName ?? "All");

            return new ConfigurationResult { IsSuccess = true, Message = "Configuration reset to defaults" };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to reset configuration to defaults");
            return new ConfigurationResult
            {
                IsSuccess = false,
                Message = "Failed to reset configuration",
                Errors = new List<string> { ex.Message }
            };
        }
    }

    public IObservable<ConfigurationChangeNotification> GetChangeNotifications()
    {
        return _changeNotifications.AsObservable();
    }

    public async Task ReloadConfigurationAsync()
    {
        try
        {
            // Clear cache to force reload
            _configurationCache.Clear();

            // Reload configuration sources
            if (_configuration is IConfigurationRoot configRoot)
            {
                configRoot.Reload();
            }

            _logger.LogInformation("Configuration reloaded successfully");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to reload configuration");
        }
    }

    public ConfigurationMetadata GetConfigurationMetadata()
    {
        return new ConfigurationMetadata
        {
            LastModified = GetLastModifiedTimes(),
            SourceFiles = GetSourceFiles(),
            ValidationResults = GetValidationResults(),
            AccessCounts = GetAccessCounts()
        };
    }

    #region Private Methods

    private T LoadConfiguration<T>(string? sectionName) where T : class, new()
    {
        try
        {
            var effectiveSectionName = sectionName ?? typeof(T).Name;

            // Try to load from user settings first, then application settings
            if (IsUserConfiguration<T>())
            {
                var userConfig = LoadFromUserSettings<T>(effectiveSectionName);
                if (userConfig != null) return userConfig;
            }

            // Load from application configuration
            var config = _configuration.GetSection(effectiveSectionName).Get<T>();
            return config ?? new T();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to load configuration for type {TypeName}", typeof(T).Name);
            return new T();
        }
    }

    private T? LoadFromUserSettings<T>(string sectionName) where T : class
    {
        try
        {
            if (!File.Exists(_userSettingsPath))
            {
                return null;
            }

            var json = File.ReadAllText(_userSettingsPath);
            var userSettings = JsonSerializer.Deserialize<JObject>(json);

            if (userSettings?.TryGetValue(sectionName, out var sectionToken) == true)
            {
                return sectionToken.ToObject<T>();
            }

            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to load user settings for section {SectionName}", sectionName);
            return null;
        }
    }

    private async Task<ValidationResult> ValidateConfigurationObjectAsync<T>(T configuration)
    {
        var context = new ValidationContext(configuration);
        var results = new List<System.ComponentModel.DataAnnotations.ValidationResult>();

        var isValid = Validator.TryValidateObject(configuration, context, results, validateAllProperties: true);

        return new ValidationResult
        {
            IsValid = isValid,
            Errors = results.Select(r => r.ErrorMessage ?? "Validation error").ToList()
        };
    }

    private string GetUserSettingsPath()
    {
        var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var bulkEditorPath = Path.Combine(appDataPath, "BulkEditor");

        Directory.CreateDirectory(bulkEditorPath);

        return Path.Combine(bulkEditorPath, "UserSettings.json");
    }

    private FileSystemWatcher SetupFileWatcher()
    {
        var watcher = new FileSystemWatcher(Path.GetDirectoryName(_userSettingsPath)!)
        {
            Filter = "UserSettings.json",
            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.CreationTime
        };

        watcher.Changed += OnUserSettingsFileChanged;
        watcher.EnableRaisingEvents = true;

        return watcher;
    }

    private void OnUserSettingsFileChanged(object sender, FileSystemEventArgs e)
    {
        // Debounce file change events and reload user settings
        Task.Delay(500).ContinueWith(_ =>
        {
            _configurationCache.Clear();
            _logger.LogInformation("User settings file changed, cache cleared");
        });
    }

    private void OnAppSettingsChanged(AppSettings appSettings, string? name)
    {
        _configurationCache.Clear();
        _logger.LogInformation("Application settings changed, cache cleared");
    }

    #endregion

    public void Dispose()
    {
        _fileWatcher?.Dispose();
        _changeNotifications?.Dispose();
        _updateSemaphore?.Dispose();
    }
}
```

## 📁 **Configuration File Structure**

### **appsettings.json** (Application Configuration)

```json
{
  "AppSettings": {
    "Api": {
      "BaseUrl": "https://api.example.com",
      "TimeoutSeconds": 30,
      "RetryCount": 3,
      "RetryDelayMs": 1000,
      "CircuitBreakerTimeoutSeconds": 30,
      "CircuitBreakerFailureThreshold": 5,
      "Headers": {
        "User-Agent": "BulkEditor/1.0",
        "Accept": "application/json"
      }
    },
    "Processing": {
      "MaxConcurrentDocuments": 5,
      "BatchSize": 20,
      "CreateBackups": true,
      "BackupRetentionDays": 30,
      "BackupLocation": "%APPDATA%\\BulkEditor\\Backups",
      "AutoCleanupBackups": true,
      "MaxBackupFiles": 50,
      "ValidateDocumentsBeforeProcessing": true,
      "CreateChangelogOnCompletion": true,
      "DefaultChangelogFormat": "Text"
    },
    "Logging": {
      "MinimumLevel": "Information",
      "LogDirectory": "%APPDATA%\\BulkEditor\\Logs",
      "FileSizeLimitMB": 50,
      "RetainedFileCountLimit": 10,
      "EnableConsoleLogging": true,
      "EnableFileLogging": true,
      "EnableStructuredLogging": true,
      "LogFilePattern": "log-{Date}.txt",
      "OutputTemplate": "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}",
      "LogLevelOverrides": {
        "Microsoft": "Warning",
        "System": "Warning"
      }
    },
    "Security": {
      "ValidateFilePaths": true,
      "AllowNetworkPaths": false,
      "RequireFileExtensionValidation": true,
      "AllowedFileExtensions": [".docx"],
      "MaxFileSizeMB": 100,
      "SanitizeUserInputs": true,
      "LogSecurityEvents": true
    },
    "Performance": {
      "MemoryCacheSizeMB": 500,
      "CacheExpirationMinutes": 15,
      "EnableMetricsCollection": true,
      "EnablePerformanceCounters": false,
      "GarbageCollectionFrequency": 3,
      "OptimizeForLargeFiles": false
    }
  }
}
```

### **UserSettings.json** (User Configuration - AppData)

```json
{
  "UI": {
    "SelectedTheme": "System",
    "AutoDetectSystemTheme": true,
    "AccentColor": "Blue",
    "WindowWidth": 1200,
    "WindowHeight": 800,
    "WindowState": "Normal",
    "RememberWindowPosition": true,
    "WindowLeft": 100,
    "WindowTop": 100,
    "ShowStatusBar": true,
    "ShowProgressDetails": true,
    "AutoRefreshDocumentList": true,
    "RecentFilesCount": 10,
    "RecentFiles": [],
    "RecentFolders": []
  },
  "Workflow": {
    "DefaultInputFolder": "",
    "DefaultOutputFolder": "",
    "AutoOpenChangelog": true,
    "AutoSaveSettings": true,
    "ConfirmBeforeProcessing": true,
    "ConfirmBeforeDeleting": true,
    "AutoSelectAllFiles": false,
    "DefaultProcessingMode": "Batch",
    "SaveProcessingResults": true,
    "ProcessingResultsFolder": "%APPDATA%\\BulkEditor\\Results",
    "CreateDetailedLogs": true,
    "FileNamingConvention": "Original"
  },
  "Notifications": {
    "ShowCompletionNotifications": true,
    "ShowErrorNotifications": true,
    "ShowWarningNotifications": true,
    "ShowProgressNotifications": false,
    "PlaySoundsOnCompletion": false,
    "ShowSystemTrayNotifications": true,
    "NotificationDurationSeconds": 5,
    "NotificationPosition": "BottomRight"
  },
  "CustomSettings": {},
  "LastModified": "2024-01-01T00:00:00Z",
  "Version": "1.0.0"
}
```

### **appsettings.Development.json** (Development Overrides)

```json
{
  "AppSettings": {
    "Api": {
      "BaseUrl": "https://dev-api.example.com",
      "TimeoutSeconds": 60
    },
    "Logging": {
      "MinimumLevel": "Debug",
      "EnableConsoleLogging": true,
      "LogLevelOverrides": {
        "BulkEditor": "Debug"
      }
    },
    "Security": {
      "LogSecurityEvents": true
    },
    "Performance": {
      "EnableMetricsCollection": true,
      "EnablePerformanceCounters": true
    }
  }
}
```

## 🔄 **Configuration Management Workflow**

### **Initialization Sequence**

1. **Load Application Settings**: appsettings.json → appsettings.{Environment}.json
2. **Load User Settings**: UserSettings.json from AppData (create if missing)
3. **Validate Configuration**: Run validation on all loaded configuration
4. **Apply Environment Variables**: Override with environment variables if present
5. **Cache Configuration**: Store validated configuration in memory cache
6. **Start File Monitoring**: Monitor configuration files for changes

### **Update Workflow**

1. **Receive Update Request**: Configuration change from UI or service
2. **Validate Changes**: Run validation on updated configuration
3. **Acquire Update Lock**: Ensure thread-safe configuration updates
4. **Update File**: Write changes to appropriate configuration file
5. **Update Cache**: Refresh in-memory cache with new values
6. **Notify Subscribers**: Send change notifications to interested parties
7. **Log Changes**: Record configuration changes for audit trail

### **Error Handling Strategy**

- **Invalid Configuration**: Fallback to defaults, log error, notify user
- **File Access Errors**: Retry with exponential backoff, use cached values
- **Validation Failures**: Reject changes, provide detailed error messages
- **Corruption Recovery**: Backup and restore from known good configuration

This configuration management system provides robust, type-safe configuration handling with comprehensive validation, persistence, and change notification capabilities.
