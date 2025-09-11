using System;
using System.Collections.Generic;

namespace BulkEditor.Core.Configuration
{
    /// <summary>
    /// Main application settings configuration
    /// </summary>
    public class AppSettings
    {
        public ProcessingSettings Processing { get; set; } = new();
        public ValidationSettings Validation { get; set; } = new();
        public BackupSettings Backup { get; set; } = new();
        public LoggingSettings Logging { get; set; } = new();
        public UiSettings UI { get; set; } = new();
        public ApiSettings Api { get; set; } = new();
        public ReplacementSettings Replacement { get; set; } = new();
        public UpdateSettings Update { get; set; } = new();
        public OfflineSettings Offline { get; set; } = new();
    }

    /// <summary>
    /// Document processing settings
    /// </summary>
    public class ProcessingSettings
    {
        public int MaxConcurrentDocuments { get; set; } = 200;
        public int BatchSize { get; set; } = 50;
        public TimeSpan TimeoutPerDocument { get; set; } = TimeSpan.FromMinutes(5);
        public bool CreateBackupBeforeProcessing { get; set; } = true;
        public bool ValidateHyperlinks { get; set; } = true;
        public bool UpdateHyperlinks { get; set; } = true;
        public bool AddContentIds { get; set; } = true;
        public bool OptimizeText { get; set; } = false;
        public List<string> SupportedExtensions { get; set; } = new() { ".docx", ".docm" };
        public string LookupIdPattern { get; set; } = @"(TSRC-[^-]+-[0-9]{6}|CMS-[^-]+-[0-9]{6})";
    }

    /// <summary>
    /// Hyperlink validation settings
    /// </summary>
    public class ValidationSettings
    {
        public TimeSpan HttpTimeout { get; set; } = TimeSpan.FromSeconds(30);
        public int MaxRetryAttempts { get; set; } = 3;
        public TimeSpan RetryDelay { get; set; } = TimeSpan.FromSeconds(2);
        public string UserAgent { get; set; } = "BulkEditor/1.0";
        public bool CheckExpiredContent { get; set; } = true;
        public bool FollowRedirects { get; set; } = true;
        public List<string> SkipDomains { get; set; } = new();
        public bool AutoReplaceTitles { get; set; } = false;
        public bool ReportTitleDifferences { get; set; } = true;
        public bool ValidateTitlesOnly { get; set; } = false;
    }

    /// <summary>
    /// Backup and recovery settings
    /// </summary>
    public class BackupSettings
    {
        public string BackupDirectory { get; set; } = "Backups";
        public bool CreateTimestampedBackups { get; set; } = true;
        public int MaxBackupAge { get; set; } = 30; // days
        public bool CompressBackups { get; set; } = false;
        public bool AutoCleanupOldBackups { get; set; } = true;
    }

    /// <summary>
    /// Logging configuration settings
    /// </summary>
    public class LoggingSettings
    {
        public string LogLevel { get; set; } = "Information";
        public string LogDirectory { get; set; } = "Logs";
        public bool EnableFileLogging { get; set; } = true;
        public bool EnableConsoleLogging { get; set; } = true;
        public int MaxLogFileSizeMB { get; set; } = 10;
        public int MaxLogFiles { get; set; } = 5;
        public string LogFilePattern { get; set; } = "bulkeditor-{Date}.log";
    }

    /// <summary>
    /// User interface settings
    /// </summary>
    public class UiSettings
    {
        public string Theme { get; set; } = "Light";
        public string Language { get; set; } = "en-US";
        public bool ShowProgressDetails { get; set; } = true;
        public bool MinimizeToSystemTray { get; set; } = false;
        public bool ConfirmBeforeProcessing { get; set; } = true;
        public bool AutoSaveSettings { get; set; } = true;
        public string ConsultantEmail { get; set; } = string.Empty; // CORRECT LOCATION
        public WindowSettings Window { get; set; } = new();
    }

    /// <summary>
    /// Window settings
    /// </summary>
    public class WindowSettings
    {
        public double Width { get; set; } = 1200;
        public double Height { get; set; } = 800;
        public double Left { get; set; } = 100;
        public double Top { get; set; } = 100;
        public bool Maximized { get; set; } = false;
    }

    /// <summary>
    /// API and web service settings
    /// </summary>
    public class ApiSettings
    {
        public string BaseUrl { get; set; } = string.Empty;
        public string ApiKey { get; set; } = string.Empty;
        public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(30);
        public bool EnableCaching { get; set; } = true;
        public TimeSpan CacheExpiry { get; set; } = TimeSpan.FromHours(1);
    }
    /// <summary>
    /// Replacement processing settings
    /// </summary>
    public class ReplacementSettings
    {
        public bool EnableHyperlinkReplacement { get; set; } = false;
        public bool EnableTextReplacement { get; set; } = false;
        public List<HyperlinkReplacementRule> HyperlinkRules { get; set; } = new();
        public List<TextReplacementRule> TextRules { get; set; } = new();
        public int MaxReplacementRules { get; set; } = 50;
        public bool ValidateContentIds { get; set; } = true;
    }

    /// <summary>
    /// Rule for hyperlink replacement
    /// </summary>
    public class HyperlinkReplacementRule
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string TitleToMatch { get; set; } = string.Empty;
        public string ContentId { get; set; } = string.Empty;
        public bool IsEnabled { get; set; } = true;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }

    /// <summary>
    /// Rule for text replacement
    /// </summary>
    public class TextReplacementRule
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string SourceText { get; set; } = string.Empty;
        public string ReplacementText { get; set; } = string.Empty;
        public bool IsEnabled { get; set; } = true;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }

    /// <summary>
    /// Auto-update settings
    /// </summary>
    public class UpdateSettings
    {
        public bool AutoUpdateEnabled { get; set; } = true;
        public int CheckIntervalHours { get; set; } = 24;
        public bool InstallSecurityUpdatesAutomatically { get; set; } = true;
        public bool NotifyOnUpdatesAvailable { get; set; } = true;
        public bool CreateBackupBeforeUpdate { get; set; } = true;
        public string GitHubOwner { get; } = "ItMeDiaTech";
        public string GitHubRepository { get; } = "Bulk_Editor_WPF";
        public DateTime LastUpdateCheck { get; set; } = DateTime.MinValue;
        public bool IncludePrerelease { get; set; } = false;
    }

    /// <summary>
    /// Offline mode and network failure handling settings
    /// </summary>
    public class OfflineSettings
    {
        public bool OfflineModeEnabled { get; set; } = false;
        public bool AllowStartupWithoutNetwork { get; set; } = true;
        public bool SkipUpdateCheckWhenOffline { get; set; } = true;
        public bool ContinueProcessingWithoutValidation { get; set; } = false;
        public int NetworkTimeoutSeconds { get; set; } = 10;
        public bool ShowOfflineIndicator { get; set; } = true;
        public bool CacheLastKnownNetworkState { get; set; } = true;
    }
}