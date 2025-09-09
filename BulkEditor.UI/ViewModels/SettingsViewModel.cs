using BulkEditor.Core.Configuration;
using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using BulkEditor.UI.ViewModels.Settings;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace BulkEditor.UI.ViewModels
{
    /// <summary>
    /// ViewModel for the Settings window
    /// </summary>
    public partial class SettingsViewModel : ViewModelBase
    {
        private readonly ILoggingService _logger;
        private readonly IConfigurationService _configurationService;
        private readonly IUpdateService _updateService;
        private readonly AppSettings _originalSettings;
        private readonly AppSettings _currentSettings;
        private readonly IHttpService _httpService;

        public event EventHandler<bool?>? RequestClose;

        public ProcessingSettingsViewModel ProcessingSettings { get; }
        public ValidationSettingsViewModel ValidationSettings { get; }
        public BackupSettingsViewModel BackupSettings { get; }
        public LoggingSettingsViewModel LoggingSettings { get; }
        public ReplacementSettingsViewModel ReplacementSettings { get; }
        public UpdateSettingsViewModel UpdateSettings { get; }

        public SettingsViewModel(AppSettings appSettings, ILoggingService logger, IConfigurationService configurationService, IUpdateService updateService, IHttpService httpService, IThemeService themeService)
        {
            _originalSettings = appSettings ?? throw new ArgumentNullException(nameof(appSettings));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _configurationService = configurationService ?? throw new ArgumentNullException(nameof(configurationService));
            _updateService = updateService ?? throw new ArgumentNullException(nameof(updateService));
            _httpService = httpService ?? throw new ArgumentNullException(nameof(httpService));

            // Create a copy of the settings to work with
            _currentSettings = CloneSettings(appSettings);

            ProcessingSettings = new ProcessingSettingsViewModel(themeService);
            ValidationSettings = new ValidationSettingsViewModel(logger, httpService);
            BackupSettings = new BackupSettingsViewModel();
            LoggingSettings = new LoggingSettingsViewModel();
            ReplacementSettings = new ReplacementSettingsViewModel();
            UpdateSettings = new UpdateSettingsViewModel(logger, updateService);

            Title = "Settings";
            LoadSettingsIntoViewModels();
        }

        private void LoadSettingsIntoViewModels()
        {
            // Processing Settings
            ProcessingSettings.MaxConcurrentDocuments = _currentSettings.Processing.MaxConcurrentDocuments;
            ProcessingSettings.BatchSize = _currentSettings.Processing.BatchSize;
            ProcessingSettings.CreateBackupBeforeProcessing = _currentSettings.Processing.CreateBackupBeforeProcessing;
            ProcessingSettings.TimeoutPerDocumentMinutes = (int)_currentSettings.Processing.TimeoutPerDocument.TotalMinutes;

            // Validation Settings
            ValidationSettings.HttpTimeoutSeconds = (int)_currentSettings.Validation.HttpTimeout.TotalSeconds;
            ValidationSettings.MaxRetryAttempts = _currentSettings.Validation.MaxRetryAttempts;
            ValidationSettings.RetryDelaySeconds = (int)_currentSettings.Validation.RetryDelay.TotalSeconds;
            ValidationSettings.UserAgent = _currentSettings.Validation.UserAgent;
            ValidationSettings.CheckExpiredContent = _currentSettings.Validation.CheckExpiredContent;
            ValidationSettings.FollowRedirects = _currentSettings.Validation.FollowRedirects;
            ValidationSettings.LookupIdPattern = _currentSettings.Processing.LookupIdPattern;
            ValidationSettings.AutoReplaceTitles = _currentSettings.Validation.AutoReplaceTitles;
            ValidationSettings.ReportTitleDifferences = _currentSettings.Validation.ReportTitleDifferences;
            ValidationSettings.ApiBaseUrl = _currentSettings.Api.BaseUrl;
            ValidationSettings.ApiKey = _currentSettings.Api.ApiKey;
            ValidationSettings.ApiTimeoutSeconds = (int)_currentSettings.Api.Timeout.TotalSeconds;
            ValidationSettings.EnableApiCaching = _currentSettings.Api.EnableCaching;
            ValidationSettings.ApiCacheExpiryHours = (int)_currentSettings.Api.CacheExpiry.TotalHours;

            // Backup Settings
            BackupSettings.BackupDirectory = _currentSettings.Backup.BackupDirectory;
            BackupSettings.CreateTimestampedBackups = _currentSettings.Backup.CreateTimestampedBackups;
            BackupSettings.CompressBackups = _currentSettings.Backup.CompressBackups;
            BackupSettings.AutoCleanupOldBackups = _currentSettings.Backup.AutoCleanupOldBackups;
            BackupSettings.MaxBackupAge = _currentSettings.Backup.MaxBackupAge;

            // Logging Settings
            LoggingSettings.LogLevel = _currentSettings.Logging.LogLevel;
            LoggingSettings.LogDirectory = _currentSettings.Logging.LogDirectory;
            LoggingSettings.EnableFileLogging = _currentSettings.Logging.EnableFileLogging;
            LoggingSettings.EnableConsoleLogging = _currentSettings.Logging.EnableConsoleLogging;
            LoggingSettings.MaxLogFileSizeMB = _currentSettings.Logging.MaxLogFileSizeMB;
            LoggingSettings.MaxLogFiles = _currentSettings.Logging.MaxLogFiles;

            // Replacement Settings
            ReplacementSettings.EnableHyperlinkReplacement = _currentSettings.Replacement.EnableHyperlinkReplacement;
            ReplacementSettings.EnableTextReplacement = _currentSettings.Replacement.EnableTextReplacement;
            ReplacementSettings.MaxReplacementRules = _currentSettings.Replacement.MaxReplacementRules;
            ReplacementSettings.ValidateContentIds = _currentSettings.Replacement.ValidateContentIds;

            // Update Settings
            UpdateSettings.AutoUpdateEnabled = _currentSettings.Update.AutoUpdateEnabled;
            UpdateSettings.CheckIntervalHours = _currentSettings.Update.CheckIntervalHours;
            UpdateSettings.InstallSecurityUpdatesAutomatically = _currentSettings.Update.InstallSecurityUpdatesAutomatically;
            UpdateSettings.NotifyOnUpdatesAvailable = _currentSettings.Update.NotifyOnUpdatesAvailable;
            UpdateSettings.CreateBackupBeforeUpdate = _currentSettings.Update.CreateBackupBeforeUpdate;
            UpdateSettings.IncludePrerelease = _currentSettings.Update.IncludePrerelease;
            UpdateSettings.GitHubOwner = _currentSettings.Update.GitHubOwner;
            UpdateSettings.GitHubRepository = _currentSettings.Update.GitHubRepository;

            // Load replacement rules into collections
            ReplacementSettings.HyperlinkRules.Clear();
            foreach (var rule in _currentSettings.Replacement.HyperlinkRules)
            {
                ReplacementSettings.HyperlinkRules.Add(rule);
            }

            ReplacementSettings.TextRules.Clear();
            foreach (var rule in _currentSettings.Replacement.TextRules)
            {
                ReplacementSettings.TextRules.Add(rule);
            }
        }

        private void UpdateSettingsFromViewModels()
        {
            // Processing Settings
            _currentSettings.Processing.MaxConcurrentDocuments = ProcessingSettings.MaxConcurrentDocuments;
            _currentSettings.Processing.BatchSize = ProcessingSettings.BatchSize;
            _currentSettings.Processing.CreateBackupBeforeProcessing = ProcessingSettings.CreateBackupBeforeProcessing;
            // Note: Processing options moved to dedicated Processing Options window
            _currentSettings.Processing.TimeoutPerDocument = TimeSpan.FromMinutes(ProcessingSettings.TimeoutPerDocumentMinutes);

            // Validation Settings
            _currentSettings.Validation.HttpTimeout = TimeSpan.FromSeconds(ValidationSettings.HttpTimeoutSeconds);
            _currentSettings.Validation.MaxRetryAttempts = ValidationSettings.MaxRetryAttempts;
            _currentSettings.Validation.RetryDelay = TimeSpan.FromSeconds(ValidationSettings.RetryDelaySeconds);
            _currentSettings.Validation.UserAgent = ValidationSettings.UserAgent;
            _currentSettings.Validation.CheckExpiredContent = ValidationSettings.CheckExpiredContent;
            _currentSettings.Validation.FollowRedirects = ValidationSettings.FollowRedirects;
            _currentSettings.Validation.AutoReplaceTitles = ValidationSettings.AutoReplaceTitles;
            _currentSettings.Validation.ReportTitleDifferences = ValidationSettings.ReportTitleDifferences;
            _currentSettings.Processing.LookupIdPattern = ValidationSettings.LookupIdPattern;

            // Backup Settings
            _currentSettings.Backup.BackupDirectory = BackupSettings.BackupDirectory;
            _currentSettings.Backup.CreateTimestampedBackups = BackupSettings.CreateTimestampedBackups;
            _currentSettings.Backup.CompressBackups = BackupSettings.CompressBackups;
            _currentSettings.Backup.AutoCleanupOldBackups = BackupSettings.AutoCleanupOldBackups;
            _currentSettings.Backup.MaxBackupAge = BackupSettings.MaxBackupAge;

            // Logging Settings
            _currentSettings.Logging.LogLevel = LoggingSettings.LogLevel;
            _currentSettings.Logging.LogDirectory = LoggingSettings.LogDirectory;
            _currentSettings.Logging.EnableFileLogging = LoggingSettings.EnableFileLogging;
            _currentSettings.Logging.EnableConsoleLogging = LoggingSettings.EnableConsoleLogging;
            _currentSettings.Logging.MaxLogFileSizeMB = LoggingSettings.MaxLogFileSizeMB;
            _currentSettings.Logging.MaxLogFiles = LoggingSettings.MaxLogFiles;

            // Replacement Settings
            _currentSettings.Replacement.EnableHyperlinkReplacement = ReplacementSettings.EnableHyperlinkReplacement;
            _currentSettings.Replacement.EnableTextReplacement = ReplacementSettings.EnableTextReplacement;
            _currentSettings.Replacement.MaxReplacementRules = ReplacementSettings.MaxReplacementRules;
            _currentSettings.Replacement.ValidateContentIds = ReplacementSettings.ValidateContentIds;

            // API Settings
            _currentSettings.Api.BaseUrl = ValidationSettings.ApiBaseUrl;
            _currentSettings.Api.ApiKey = ValidationSettings.ApiKey;
            _currentSettings.Api.Timeout = TimeSpan.FromSeconds(ValidationSettings.ApiTimeoutSeconds);
            _currentSettings.Api.EnableCaching = ValidationSettings.EnableApiCaching;
            _currentSettings.Api.CacheExpiry = TimeSpan.FromHours(ValidationSettings.ApiCacheExpiryHours);

            // Update Settings
            _currentSettings.Update.AutoUpdateEnabled = UpdateSettings.AutoUpdateEnabled;
            _currentSettings.Update.CheckIntervalHours = UpdateSettings.CheckIntervalHours;
            _currentSettings.Update.InstallSecurityUpdatesAutomatically = UpdateSettings.InstallSecurityUpdatesAutomatically;
            _currentSettings.Update.NotifyOnUpdatesAvailable = UpdateSettings.NotifyOnUpdatesAvailable;
            _currentSettings.Update.CreateBackupBeforeUpdate = UpdateSettings.CreateBackupBeforeUpdate;
            _currentSettings.Update.IncludePrerelease = UpdateSettings.IncludePrerelease;
            _currentSettings.Update.GitHubOwner = UpdateSettings.GitHubOwner;
            _currentSettings.Update.GitHubRepository = UpdateSettings.GitHubRepository;

            // Update replacement rules from collections
            _currentSettings.Replacement.HyperlinkRules.Clear();
            _currentSettings.Replacement.HyperlinkRules.AddRange(ReplacementSettings.HyperlinkRules);

            _currentSettings.Replacement.TextRules.Clear();
            _currentSettings.Replacement.TextRules.AddRange(ReplacementSettings.TextRules);
        }

        [RelayCommand]
        private void BrowseBackupDirectory()
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Backup Directory",
                CheckFileExists = false,
                CheckPathExists = true,
                FileName = "Folder Selection.",
                Filter = "Folders|*.none"
            };

            if (dialog.ShowDialog() == true)
            {
                BackupSettings.BackupDirectory = Path.GetDirectoryName(dialog.FileName) ?? "";
            }
        }

        [RelayCommand]
        private void BrowseLogDirectory()
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Log Directory",
                CheckFileExists = false,
                CheckPathExists = true,
                FileName = "Folder Selection.",
                Filter = "Folders|*.none"
            };

            if (dialog.ShowDialog() == true)
            {
                LoggingSettings.LogDirectory = Path.GetDirectoryName(dialog.FileName) ?? "";
            }
        }

        [RelayCommand]
        private void ResetToDefaults()
        {
            try
            {
                var defaultSettings = CreateDefaultSettings();

                // Copy default settings to current settings
                _currentSettings.Processing = defaultSettings.Processing;
                _currentSettings.Validation = defaultSettings.Validation;
                _currentSettings.Backup = defaultSettings.Backup;
                _currentSettings.Logging = defaultSettings.Logging;
                _currentSettings.Replacement = defaultSettings.Replacement;
                _currentSettings.Api = defaultSettings.Api;
                _currentSettings.Update = defaultSettings.Update;

                // Reload properties from updated settings - this will trigger PropertyChanged
                LoadSettingsIntoViewModels();

                _logger.LogInformation("Settings reset to defaults");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error resetting settings to defaults");
            }
        }

        [RelayCommand]
        private void Cancel()
        {
            _logger.LogInformation("Settings dialog cancelled");
            RequestClose?.Invoke(this, false);
        }

        [RelayCommand]
        private async Task SaveSettings()
        {
            try
            {
                IsBusy = true;
                BusyMessage = "Saving settings...";

                _logger.LogInformation("Starting save settings process");

                // Update current settings from UI properties
                UpdateSettingsFromViewModels();
                _logger.LogInformation("Settings updated from properties");

                // Validate settings
                if (!ValidateSettings())
                {
                    _logger.LogWarning("Settings validation failed, save aborted");
                    IsBusy = false;
                    return;
                }
                _logger.LogInformation("Settings validation passed");

                // Copy current settings back to original
                CopySettingsTo(_currentSettings, _originalSettings);
                _logger.LogInformation("Settings copied to original");

                // Ensure AppData directories exist before saving
                await _configurationService.InitializeAsync();
                _logger.LogInformation("Configuration service initialized");

                // Save settings to file
                await SaveSettingsToFileAsync();
                _logger.LogInformation("Settings saved to file successfully");

                IsBusy = false;
                RequestClose?.Invoke(this, true);
                _logger.LogInformation("Settings dialog closed with success result");
            }
            catch (Exception ex)
            {
                IsBusy = false;
                _logger.LogError(ex, "Error saving settings: {Message}", ex.Message);
                _logger.LogError("Stack trace: {StackTrace}", ex.StackTrace ?? "No stack trace available");
                // Re-throw to ensure the error is properly handled by the UI
                throw;
            }
        }

        private bool ValidateSettings()
        {
            // Validate processing settings
            if (ProcessingSettings.MaxConcurrentDocuments < 1 || ProcessingSettings.MaxConcurrentDocuments > 100)
            {
                _logger.LogWarning("Max concurrent documents must be between 1 and 100");
                return false;
            }

            if (ProcessingSettings.BatchSize < 1 || ProcessingSettings.BatchSize > 1000)
            {
                _logger.LogWarning("Batch size must be between 1 and 1000");
                return false;
            }

            if (ProcessingSettings.TimeoutPerDocumentMinutes < 1 || ProcessingSettings.TimeoutPerDocumentMinutes > 60)
            {
                _logger.LogWarning("Timeout per document must be between 1 and 60 minutes");
                return false;
            }

            // Validate validation settings
            if (ValidationSettings.HttpTimeoutSeconds < 1 || ValidationSettings.HttpTimeoutSeconds > 300)
            {
                _logger.LogWarning("HTTP timeout must be between 1 and 300 seconds");
                return false;
            }

            if (ValidationSettings.MaxRetryAttempts < 0 || ValidationSettings.MaxRetryAttempts > 10)
            {
                _logger.LogWarning("Max retry attempts must be between 0 and 10");
                return false;
            }

            // Validate backup settings
            if (BackupSettings.MaxBackupAge < 1 || BackupSettings.MaxBackupAge > 365)
            {
                _logger.LogWarning("Max backup age must be between 1 and 365 days");
                return false;
            }

            // Validate logging settings
            if (LoggingSettings.MaxLogFileSizeMB < 1 || LoggingSettings.MaxLogFileSizeMB > 1000)
            {
                _logger.LogWarning("Max log file size must be between 1 and 1000 MB");
                return false;
            }

            if (LoggingSettings.MaxLogFiles < 1 || LoggingSettings.MaxLogFiles > 100)
            {
                _logger.LogWarning("Max log files must be between 1 and 100");
                return false;
            }

            // Validate replacement settings
            if (ReplacementSettings.MaxReplacementRules < 1 || ReplacementSettings.MaxReplacementRules > 1000)
            {
                _logger.LogWarning("Max replacement rules must be between 1 and 1000");
                return false;
            }

            // Validate API settings
            if (ValidationSettings.ApiTimeoutSeconds < 1 || ValidationSettings.ApiTimeoutSeconds > 300)
            {
                _logger.LogWarning("API timeout must be between 1 and 300 seconds");
                return false;
            }

            if (ValidationSettings.ApiCacheExpiryHours < 1 || ValidationSettings.ApiCacheExpiryHours > 24)
            {
                _logger.LogWarning("API cache expiry must be between 1 and 24 hours");
                return false;
            }

            // Validate API URL format if provided
            if (!string.IsNullOrWhiteSpace(ValidationSettings.ApiBaseUrl) && ValidationSettings.ApiBaseUrl.ToLower() != "test")
            {
                if (!Uri.TryCreate(ValidationSettings.ApiBaseUrl, UriKind.Absolute, out var uri) ||
                    (uri.Scheme != "http" && uri.Scheme != "https"))
                {
                    _logger.LogWarning("API Base URL must be a valid HTTP or HTTPS URL");
                    return false;
                }
            }

            // Validate hyperlink replacement rules if enabled
            if (ReplacementSettings.EnableHyperlinkReplacement)
            {
                var validHyperlinkRules = ReplacementSettings.HyperlinkRules.Where(r =>
                    !string.IsNullOrWhiteSpace(r.TitleToMatch) &&
                    !string.IsNullOrWhiteSpace(r.ContentId)).ToList();

                foreach (var rule in validHyperlinkRules)
                {
                    // Validate Content ID format (should contain 6 digits)
                    if (!System.Text.RegularExpressions.Regex.IsMatch(rule.ContentId, @"[0-9]{6}"))
                    {
                        _logger.LogWarning("Hyperlink rule has invalid Content ID format: {ContentId}", rule.ContentId);
                        return false;
                    }
                }
            }

            // Validate text replacement rules if enabled
            if (ReplacementSettings.EnableTextReplacement)
            {
                var validTextRules = ReplacementSettings.TextRules.Where(r =>
                    !string.IsNullOrWhiteSpace(r.SourceText) &&
                    !string.IsNullOrWhiteSpace(r.ReplacementText)).ToList();

                foreach (var rule in validTextRules)
                {
                    // Check for potential infinite loops (source = replacement)
                    if (rule.SourceText.Trim().Equals(rule.ReplacementText.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        _logger.LogWarning("Text rule has identical source and replacement text: {SourceText}", rule.SourceText);
                        return false;
                    }
                }
            }

            // Validate update settings
            if (UpdateSettings.CheckIntervalHours < 1 || UpdateSettings.CheckIntervalHours > 168) // 1 hour to 1 week
            {
                _logger.LogWarning("Check interval hours must be between 1 and 168 (1 week)");
                return false;
            }

            // Validate GitHub repository settings if auto-update is enabled
            if (UpdateSettings.AutoUpdateEnabled)
            {
                if (string.IsNullOrWhiteSpace(UpdateSettings.GitHubOwner))
                {
                    _logger.LogWarning("GitHub owner is required when auto-update is enabled");
                    return false;
                }

                if (string.IsNullOrWhiteSpace(UpdateSettings.GitHubRepository))
                {
                    _logger.LogWarning("GitHub repository is required when auto-update is enabled");
                    return false;
                }
            }

            return true;
        }

        private async Task SaveSettingsToFileAsync()
        {
            try
            {
                await _configurationService.SaveSettingsAsync(_currentSettings);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving settings to file");
                throw;
            }
        }

        private static AppSettings CloneSettings(AppSettings source)
        {
            // Simple deep clone using JSON serialization
            var json = System.Text.Json.JsonSerializer.Serialize(source);
            return System.Text.Json.JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
        }

        private static void CopySettingsTo(AppSettings source, AppSettings target)
        {
            target.Processing = source.Processing;
            target.Validation = source.Validation;
            target.Backup = source.Backup;
            target.Logging = source.Logging;
            target.Replacement = source.Replacement;
            target.Api = source.Api;
            target.Update = source.Update;
        }

        private static AppSettings CreateDefaultSettings()
        {
            return new AppSettings
            {
                Processing = new ProcessingSettings
                {
                    MaxConcurrentDocuments = 10,
                    BatchSize = 50,
                    TimeoutPerDocument = TimeSpan.FromMinutes(5),
                    CreateBackupBeforeProcessing = true,
                    // Processing options now managed in dedicated Processing Options window
                    LookupIdPattern = @"(TSRC-[^-]+-[0-9]{6}|CMS-[^-]+-[0-9]{6})"
                },
                Validation = new ValidationSettings
                {
                    HttpTimeout = TimeSpan.FromSeconds(30),
                    MaxRetryAttempts = 3,
                    RetryDelay = TimeSpan.FromSeconds(2),
                    UserAgent = "BulkEditor/1.0",
                    CheckExpiredContent = true,
                    FollowRedirects = true,
                    AutoReplaceTitles = false,
                    ReportTitleDifferences = true
                },
                Backup = new BackupSettings
                {
                    BackupDirectory = "Backups",
                    CreateTimestampedBackups = true,
                    MaxBackupAge = 30,
                    CompressBackups = false,
                    AutoCleanupOldBackups = true
                },
                Logging = new LoggingSettings
                {
                    LogLevel = "Information",
                    LogDirectory = "Logs",
                    EnableFileLogging = true,
                    EnableConsoleLogging = true,
                    MaxLogFileSizeMB = 10,
                    MaxLogFiles = 5,
                    LogFilePattern = "bulkeditor-{Date}.log"
                },
                Replacement = new ReplacementSettings
                {
                    EnableHyperlinkReplacement = false,
                    EnableTextReplacement = false,
                    MaxReplacementRules = 50,
                    ValidateContentIds = true
                },
                Api = new ApiSettings
                {
                    BaseUrl = string.Empty,
                    ApiKey = string.Empty,
                    Timeout = TimeSpan.FromSeconds(30),
                    EnableCaching = true,
                    CacheExpiry = TimeSpan.FromHours(1)
                },
                Update = new UpdateSettings
                {
                    AutoUpdateEnabled = true,
                    CheckIntervalHours = 24,
                    InstallSecurityUpdatesAutomatically = true,
                    NotifyOnUpdatesAvailable = true,
                    CreateBackupBeforeUpdate = true,
                    GitHubOwner = "DiaTech",
                    GitHubRepository = "Bulk_Editor",
                    IncludePrerelease = false
                }
            };
        }



        [RelayCommand]
        private void OpenLogsFolder()
        {
            try
            {
                var logDirectory = _currentSettings.Logging.LogDirectory;

                if (!Directory.Exists(logDirectory))
                {
                    _logger.LogWarning("Log directory does not exist: {LogDirectory}", logDirectory);
                    return;
                }

                // Open the logs folder in Windows Explorer
                var startInfo = new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"\"{logDirectory}\"",
                    UseShellExecute = true
                };

                Process.Start(startInfo);
                _logger.LogInformation("Opened logs folder: {LogDirectory}", logDirectory);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to open logs folder");
            }
        }

    }
}