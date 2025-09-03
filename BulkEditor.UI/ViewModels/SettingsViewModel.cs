using BulkEditor.Core.Configuration;
using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
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

        public event EventHandler<bool?>? RequestClose;

        // Processing Settings
        [ObservableProperty]
        private int _maxConcurrentDocuments;

        [ObservableProperty]
        private int _batchSize;

        [ObservableProperty]
        private bool _createBackupBeforeProcessing;

        [ObservableProperty]
        private bool _validateHyperlinks;

        [ObservableProperty]
        private bool _updateHyperlinks;

        [ObservableProperty]
        private bool _addContentIds;

        [ObservableProperty]
        private bool _optimizeText;

        [ObservableProperty]
        private int _timeoutPerDocumentMinutes;

        // Validation Settings
        [ObservableProperty]
        private int _httpTimeoutSeconds;

        [ObservableProperty]
        private int _maxRetryAttempts;

        [ObservableProperty]
        private int _retryDelaySeconds;

        [ObservableProperty]
        private string _userAgent = string.Empty;

        [ObservableProperty]
        private bool _checkExpiredContent;

        [ObservableProperty]
        private bool _followRedirects;

        [ObservableProperty]
        private string _lookupIdPattern = string.Empty;

        [ObservableProperty]
        private bool _autoReplaceTitles;

        [ObservableProperty]
        private bool _reportTitleDifferences;

        // API Settings
        [ObservableProperty]
        private string _apiBaseUrl = string.Empty;

        [ObservableProperty]
        private string _apiKey = string.Empty;

        [ObservableProperty]
        private int _apiTimeoutSeconds;

        [ObservableProperty]
        private bool _enableApiCaching;

        [ObservableProperty]
        private int _apiCacheExpiryHours;

        // Backup Settings
        [ObservableProperty]
        private string _backupDirectory = string.Empty;

        [ObservableProperty]
        private bool _createTimestampedBackups;

        [ObservableProperty]
        private bool _compressBackups;

        [ObservableProperty]
        private bool _autoCleanupOldBackups;

        [ObservableProperty]
        private int _maxBackupAge;

        // Logging Settings
        [ObservableProperty]
        private string _logLevel = string.Empty;

        [ObservableProperty]
        private string _logDirectory = string.Empty;

        [ObservableProperty]
        private bool _enableFileLogging;

        [ObservableProperty]
        private bool _enableConsoleLogging;

        [ObservableProperty]
        private int _maxLogFileSizeMB;

        [ObservableProperty]
        private int _maxLogFiles;

        // Replacement Settings
        [ObservableProperty]
        private bool _enableHyperlinkReplacement;

        [ObservableProperty]
        private bool _enableTextReplacement;

        [ObservableProperty]
        private int _maxReplacementRules;

        [ObservableProperty]
        private bool _validateContentIds;

        // Update Settings
        [ObservableProperty]
        private bool _autoUpdateEnabled;

        [ObservableProperty]
        private int _checkIntervalHours;

        [ObservableProperty]
        private bool _installSecurityUpdatesAutomatically;

        [ObservableProperty]
        private bool _notifyOnUpdatesAvailable;

        [ObservableProperty]
        private bool _createBackupBeforeUpdate;

        [ObservableProperty]
        private bool _includePrerelease;

        [ObservableProperty]
        private string _gitHubOwner = string.Empty;

        [ObservableProperty]
        private string _gitHubRepository = string.Empty;

        // Version Information
        [ObservableProperty]
        private string _currentVersion = string.Empty;

        [ObservableProperty]
        private string _latestVersion = string.Empty;

        [ObservableProperty]
        private bool _updateAvailable = false;

        [ObservableProperty]
        private string _updateCheckStatus = "Click 'Check for Updates' to check for the latest version";

        [ObservableProperty]
        private string _releaseNotes = string.Empty;

        // Replacement Rules Collections
        public ObservableCollection<HyperlinkReplacementRule> HyperlinkRules { get; set; } = new();
        public ObservableCollection<TextReplacementRule> TextRules { get; set; } = new();

        public SettingsViewModel(AppSettings appSettings, ILoggingService logger, IConfigurationService configurationService, IUpdateService updateService)
        {
            _originalSettings = appSettings ?? throw new ArgumentNullException(nameof(appSettings));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _configurationService = configurationService ?? throw new ArgumentNullException(nameof(configurationService));
            _updateService = updateService ?? throw new ArgumentNullException(nameof(updateService));

            // Create a copy of the settings to work with
            _currentSettings = CloneSettings(appSettings);

            Title = "Settings";
            LoadSettingsIntoProperties();
            LoadVersionInformation();
        }

        private void LoadSettingsIntoProperties()
        {
            // Processing Settings
            MaxConcurrentDocuments = _currentSettings.Processing.MaxConcurrentDocuments;
            BatchSize = _currentSettings.Processing.BatchSize;
            CreateBackupBeforeProcessing = _currentSettings.Processing.CreateBackupBeforeProcessing;
            ValidateHyperlinks = _currentSettings.Processing.ValidateHyperlinks;
            UpdateHyperlinks = _currentSettings.Processing.UpdateHyperlinks;
            AddContentIds = _currentSettings.Processing.AddContentIds;
            OptimizeText = _currentSettings.Processing.OptimizeText;
            TimeoutPerDocumentMinutes = (int)_currentSettings.Processing.TimeoutPerDocument.TotalMinutes;

            // Validation Settings
            HttpTimeoutSeconds = (int)_currentSettings.Validation.HttpTimeout.TotalSeconds;
            MaxRetryAttempts = _currentSettings.Validation.MaxRetryAttempts;
            RetryDelaySeconds = (int)_currentSettings.Validation.RetryDelay.TotalSeconds;
            UserAgent = _currentSettings.Validation.UserAgent;
            CheckExpiredContent = _currentSettings.Validation.CheckExpiredContent;
            FollowRedirects = _currentSettings.Validation.FollowRedirects;
            LookupIdPattern = _currentSettings.Processing.LookupIdPattern;

            // Backup Settings
            BackupDirectory = _currentSettings.Backup.BackupDirectory;
            CreateTimestampedBackups = _currentSettings.Backup.CreateTimestampedBackups;
            CompressBackups = _currentSettings.Backup.CompressBackups;
            AutoCleanupOldBackups = _currentSettings.Backup.AutoCleanupOldBackups;
            MaxBackupAge = _currentSettings.Backup.MaxBackupAge;

            // Logging Settings
            LogLevel = _currentSettings.Logging.LogLevel;
            LogDirectory = _currentSettings.Logging.LogDirectory;
            EnableFileLogging = _currentSettings.Logging.EnableFileLogging;
            EnableConsoleLogging = _currentSettings.Logging.EnableConsoleLogging;
            MaxLogFileSizeMB = _currentSettings.Logging.MaxLogFileSizeMB;
            MaxLogFiles = _currentSettings.Logging.MaxLogFiles;

            // Replacement Settings
            EnableHyperlinkReplacement = _currentSettings.Replacement.EnableHyperlinkReplacement;
            EnableTextReplacement = _currentSettings.Replacement.EnableTextReplacement;
            MaxReplacementRules = _currentSettings.Replacement.MaxReplacementRules;
            ValidateContentIds = _currentSettings.Replacement.ValidateContentIds;

            // Title replacement settings
            AutoReplaceTitles = _currentSettings.Validation.AutoReplaceTitles;
            ReportTitleDifferences = _currentSettings.Validation.ReportTitleDifferences;

            // API Settings
            ApiBaseUrl = _currentSettings.Api.BaseUrl;
            ApiKey = _currentSettings.Api.ApiKey;
            ApiTimeoutSeconds = (int)_currentSettings.Api.Timeout.TotalSeconds;
            EnableApiCaching = _currentSettings.Api.EnableCaching;
            ApiCacheExpiryHours = (int)_currentSettings.Api.CacheExpiry.TotalHours;

            // Update Settings
            AutoUpdateEnabled = _currentSettings.Update.AutoUpdateEnabled;
            CheckIntervalHours = _currentSettings.Update.CheckIntervalHours;
            InstallSecurityUpdatesAutomatically = _currentSettings.Update.InstallSecurityUpdatesAutomatically;
            NotifyOnUpdatesAvailable = _currentSettings.Update.NotifyOnUpdatesAvailable;
            CreateBackupBeforeUpdate = _currentSettings.Update.CreateBackupBeforeUpdate;
            IncludePrerelease = _currentSettings.Update.IncludePrerelease;
            GitHubOwner = _currentSettings.Update.GitHubOwner;
            GitHubRepository = _currentSettings.Update.GitHubRepository;

            // Load replacement rules into collections
            HyperlinkRules.Clear();
            foreach (var rule in _currentSettings.Replacement.HyperlinkRules)
            {
                HyperlinkRules.Add(rule);
            }

            TextRules.Clear();
            foreach (var rule in _currentSettings.Replacement.TextRules)
            {
                TextRules.Add(rule);
            }
        }

        private void UpdateSettingsFromProperties()
        {
            // Processing Settings
            _currentSettings.Processing.MaxConcurrentDocuments = MaxConcurrentDocuments;
            _currentSettings.Processing.BatchSize = BatchSize;
            _currentSettings.Processing.CreateBackupBeforeProcessing = CreateBackupBeforeProcessing;
            _currentSettings.Processing.ValidateHyperlinks = ValidateHyperlinks;
            _currentSettings.Processing.UpdateHyperlinks = UpdateHyperlinks;
            _currentSettings.Processing.AddContentIds = AddContentIds;
            _currentSettings.Processing.OptimizeText = OptimizeText;
            _currentSettings.Processing.TimeoutPerDocument = TimeSpan.FromMinutes(TimeoutPerDocumentMinutes);

            // Validation Settings
            _currentSettings.Validation.HttpTimeout = TimeSpan.FromSeconds(HttpTimeoutSeconds);
            _currentSettings.Validation.MaxRetryAttempts = MaxRetryAttempts;
            _currentSettings.Validation.RetryDelay = TimeSpan.FromSeconds(RetryDelaySeconds);
            _currentSettings.Validation.UserAgent = UserAgent;
            _currentSettings.Validation.CheckExpiredContent = CheckExpiredContent;
            _currentSettings.Validation.FollowRedirects = FollowRedirects;
            _currentSettings.Validation.AutoReplaceTitles = AutoReplaceTitles;
            _currentSettings.Validation.ReportTitleDifferences = ReportTitleDifferences;
            _currentSettings.Processing.LookupIdPattern = LookupIdPattern;

            // Backup Settings
            _currentSettings.Backup.BackupDirectory = BackupDirectory;
            _currentSettings.Backup.CreateTimestampedBackups = CreateTimestampedBackups;
            _currentSettings.Backup.CompressBackups = CompressBackups;
            _currentSettings.Backup.AutoCleanupOldBackups = AutoCleanupOldBackups;
            _currentSettings.Backup.MaxBackupAge = MaxBackupAge;

            // Logging Settings
            _currentSettings.Logging.LogLevel = LogLevel;
            _currentSettings.Logging.LogDirectory = LogDirectory;
            _currentSettings.Logging.EnableFileLogging = EnableFileLogging;
            _currentSettings.Logging.EnableConsoleLogging = EnableConsoleLogging;
            _currentSettings.Logging.MaxLogFileSizeMB = MaxLogFileSizeMB;
            _currentSettings.Logging.MaxLogFiles = MaxLogFiles;

            // Replacement Settings
            _currentSettings.Replacement.EnableHyperlinkReplacement = EnableHyperlinkReplacement;
            _currentSettings.Replacement.EnableTextReplacement = EnableTextReplacement;
            _currentSettings.Replacement.MaxReplacementRules = MaxReplacementRules;
            _currentSettings.Replacement.ValidateContentIds = ValidateContentIds;

            // Title replacement settings
            _currentSettings.Validation.AutoReplaceTitles = AutoReplaceTitles;
            _currentSettings.Validation.ReportTitleDifferences = ReportTitleDifferences;

            // API Settings
            _currentSettings.Api.BaseUrl = ApiBaseUrl;
            _currentSettings.Api.ApiKey = ApiKey;
            _currentSettings.Api.Timeout = TimeSpan.FromSeconds(ApiTimeoutSeconds);
            _currentSettings.Api.EnableCaching = EnableApiCaching;
            _currentSettings.Api.CacheExpiry = TimeSpan.FromHours(ApiCacheExpiryHours);

            // Update Settings
            _currentSettings.Update.AutoUpdateEnabled = AutoUpdateEnabled;
            _currentSettings.Update.CheckIntervalHours = CheckIntervalHours;
            _currentSettings.Update.InstallSecurityUpdatesAutomatically = InstallSecurityUpdatesAutomatically;
            _currentSettings.Update.NotifyOnUpdatesAvailable = NotifyOnUpdatesAvailable;
            _currentSettings.Update.CreateBackupBeforeUpdate = CreateBackupBeforeUpdate;
            _currentSettings.Update.IncludePrerelease = IncludePrerelease;
            _currentSettings.Update.GitHubOwner = GitHubOwner;
            _currentSettings.Update.GitHubRepository = GitHubRepository;

            // Update replacement rules from collections
            _currentSettings.Replacement.HyperlinkRules.Clear();
            _currentSettings.Replacement.HyperlinkRules.AddRange(HyperlinkRules);

            _currentSettings.Replacement.TextRules.Clear();
            _currentSettings.Replacement.TextRules.AddRange(TextRules);
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
                BackupDirectory = Path.GetDirectoryName(dialog.FileName) ?? "";
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
                LogDirectory = Path.GetDirectoryName(dialog.FileName) ?? "";
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
                LoadSettingsIntoProperties();

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
                UpdateSettingsFromProperties();
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
                _logger.LogError("Stack trace: {StackTrace}", ex.StackTrace);
                // Re-throw to ensure the error is properly handled by the UI
                throw;
            }
        }

        private bool ValidateSettings()
        {
            // Validate processing settings
            if (MaxConcurrentDocuments < 1 || MaxConcurrentDocuments > 100)
            {
                _logger.LogWarning("Max concurrent documents must be between 1 and 100");
                return false;
            }

            if (BatchSize < 1 || BatchSize > 1000)
            {
                _logger.LogWarning("Batch size must be between 1 and 1000");
                return false;
            }

            if (TimeoutPerDocumentMinutes < 1 || TimeoutPerDocumentMinutes > 60)
            {
                _logger.LogWarning("Timeout per document must be between 1 and 60 minutes");
                return false;
            }

            // Validate validation settings
            if (HttpTimeoutSeconds < 1 || HttpTimeoutSeconds > 300)
            {
                _logger.LogWarning("HTTP timeout must be between 1 and 300 seconds");
                return false;
            }

            if (MaxRetryAttempts < 0 || MaxRetryAttempts > 10)
            {
                _logger.LogWarning("Max retry attempts must be between 0 and 10");
                return false;
            }

            // Validate backup settings
            if (MaxBackupAge < 1 || MaxBackupAge > 365)
            {
                _logger.LogWarning("Max backup age must be between 1 and 365 days");
                return false;
            }

            // Validate logging settings
            if (MaxLogFileSizeMB < 1 || MaxLogFileSizeMB > 1000)
            {
                _logger.LogWarning("Max log file size must be between 1 and 1000 MB");
                return false;
            }

            if (MaxLogFiles < 1 || MaxLogFiles > 100)
            {
                _logger.LogWarning("Max log files must be between 1 and 100");
                return false;
            }

            // Validate replacement settings
            if (MaxReplacementRules < 1 || MaxReplacementRules > 1000)
            {
                _logger.LogWarning("Max replacement rules must be between 1 and 1000");
                return false;
            }

            // Validate API settings
            if (ApiTimeoutSeconds < 1 || ApiTimeoutSeconds > 300)
            {
                _logger.LogWarning("API timeout must be between 1 and 300 seconds");
                return false;
            }

            if (ApiCacheExpiryHours < 1 || ApiCacheExpiryHours > 24)
            {
                _logger.LogWarning("API cache expiry must be between 1 and 24 hours");
                return false;
            }

            // Validate API URL format if provided
            if (!string.IsNullOrWhiteSpace(ApiBaseUrl) && ApiBaseUrl.ToLower() != "test")
            {
                if (!Uri.TryCreate(ApiBaseUrl, UriKind.Absolute, out var uri) ||
                    (uri.Scheme != "http" && uri.Scheme != "https"))
                {
                    _logger.LogWarning("API Base URL must be a valid HTTP or HTTPS URL");
                    return false;
                }
            }

            // Validate hyperlink replacement rules if enabled
            if (EnableHyperlinkReplacement)
            {
                var validHyperlinkRules = HyperlinkRules.Where(r =>
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
            if (EnableTextReplacement)
            {
                var validTextRules = TextRules.Where(r =>
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
            if (CheckIntervalHours < 1 || CheckIntervalHours > 168) // 1 hour to 1 week
            {
                _logger.LogWarning("Check interval hours must be between 1 and 168 (1 week)");
                return false;
            }

            // Validate GitHub repository settings if auto-update is enabled
            if (AutoUpdateEnabled)
            {
                if (string.IsNullOrWhiteSpace(GitHubOwner))
                {
                    _logger.LogWarning("GitHub owner is required when auto-update is enabled");
                    return false;
                }

                if (string.IsNullOrWhiteSpace(GitHubRepository))
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
                    ValidateHyperlinks = true,
                    UpdateHyperlinks = true,
                    AddContentIds = true,
                    OptimizeText = false,
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

        // Replacement Rule Management Commands
        [RelayCommand]
        private void AddHyperlinkRule()
        {
            if (HyperlinkRules.Count >= MaxReplacementRules)
            {
                _logger.LogWarning("Maximum number of hyperlink replacement rules reached: {MaxRules}", MaxReplacementRules);
                return;
            }

            HyperlinkRules.Add(new HyperlinkReplacementRule());
            _logger.LogInformation("Added new hyperlink replacement rule");
        }

        [RelayCommand]
        private void RemoveHyperlinkRule(HyperlinkReplacementRule rule)
        {
            if (rule != null)
            {
                HyperlinkRules.Remove(rule);
                _logger.LogInformation("Removed hyperlink replacement rule: {RuleId}", rule.Id);
            }
        }

        [RelayCommand]
        private void AddTextRule()
        {
            if (TextRules.Count >= MaxReplacementRules)
            {
                _logger.LogWarning("Maximum number of text replacement rules reached: {MaxRules}", MaxReplacementRules);
                return;
            }

            TextRules.Add(new TextReplacementRule());
            _logger.LogInformation("Added new text replacement rule");
        }

        [RelayCommand]
        private void RemoveTextRule(TextReplacementRule rule)
        {
            if (rule != null)
            {
                TextRules.Remove(rule);
                _logger.LogInformation("Removed text replacement rule: {RuleId}", rule.Id);
            }
        }

        [RelayCommand]
        private void ClearInvalidRules()
        {
            // Remove hyperlink rules with blank fields
            var invalidHyperlinkRules = HyperlinkRules.Where(r =>
                string.IsNullOrWhiteSpace(r.TitleToMatch) ||
                string.IsNullOrWhiteSpace(r.ContentId)).ToList();

            foreach (var rule in invalidHyperlinkRules)
            {
                HyperlinkRules.Remove(rule);
            }

            // Remove text rules with blank fields
            var invalidTextRules = TextRules.Where(r =>
                string.IsNullOrWhiteSpace(r.SourceText) ||
                string.IsNullOrWhiteSpace(r.ReplacementText)).ToList();

            foreach (var rule in invalidTextRules)
            {
                TextRules.Remove(rule);
            }

            if (invalidHyperlinkRules.Any() || invalidTextRules.Any())
            {
                _logger.LogInformation("Cleared {HyperlinkCount} invalid hyperlink rules and {TextCount} invalid text rules",
                    invalidHyperlinkRules.Count, invalidTextRules.Count);
            }
        }

        [RelayCommand]
        private async Task TestApiConnection()
        {
            try
            {
                IsBusy = true;
                BusyMessage = "Testing API connection...";

                if (string.IsNullOrWhiteSpace(ApiBaseUrl))
                {
                    _logger.LogWarning("API Base URL is required for connection test");
                    System.Windows.MessageBox.Show(
                        "API Base URL is required for connection test. Please enter a valid URL.",
                        "API Test Failed",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                    return;
                }

                if (ApiBaseUrl.ToLower() == "test")
                {
                    _logger.LogInformation("API connection test successful (Test mode)");
                    System.Windows.MessageBox.Show(
                        "API connection test successful!\n\nTest mode is active - using mock API responses.",
                        "API Test Successful",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                    return;
                }

                // Test actual API connection
                using var httpClient = new System.Net.Http.HttpClient();
                httpClient.Timeout = TimeSpan.FromSeconds(Math.Max(ApiTimeoutSeconds, 5)); // Minimum 5 seconds

                // Add API key if provided
                if (!string.IsNullOrWhiteSpace(ApiKey))
                {
                    httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {ApiKey}");
                }

                _logger.LogInformation("Testing API connection to: {ApiUrl}", ApiBaseUrl);
                var response = await httpClient.GetAsync(ApiBaseUrl);

                if (response.IsSuccessStatusCode)
                {
                    var responseTime = httpClient.Timeout.TotalSeconds;
                    _logger.LogInformation("API connection test successful - Status: {StatusCode}", response.StatusCode);

                    System.Windows.MessageBox.Show(
                        $"API connection test successful!\n\n" +
                        $"URL: {ApiBaseUrl}\n" +
                        $"Status: {response.StatusCode} {response.ReasonPhrase}\n" +
                        $"Response Time: < {ApiTimeoutSeconds}s\n" +
                        $"API Key: {(string.IsNullOrWhiteSpace(ApiKey) ? "Not provided" : "Configured")}",
                        "API Test Successful",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                }
                else
                {
                    _logger.LogWarning("API connection test failed - Status: {StatusCode} {ReasonPhrase}", response.StatusCode, response.ReasonPhrase);

                    System.Windows.MessageBox.Show(
                        $"API connection test failed!\n\n" +
                        $"URL: {ApiBaseUrl}\n" +
                        $"Status: {response.StatusCode} {response.ReasonPhrase}\n" +
                        $"Please check the URL and your internet connection.\n\n" +
                        $"If this is a private API, ensure your API key is correct.",
                        "API Test Failed",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Error);
                }
            }
            catch (TaskCanceledException ex) when (ex.InnerException is TimeoutException)
            {
                _logger.LogWarning("API connection test timed out after {TimeoutSeconds} seconds", ApiTimeoutSeconds);

                System.Windows.MessageBox.Show(
                    $"API connection test timed out!\n\n" +
                    $"URL: {ApiBaseUrl}\n" +
                    $"Timeout: {ApiTimeoutSeconds} seconds\n\n" +
                    $"The API may be slow or unreachable. Try increasing the timeout value or check your internet connection.",
                    "API Test Timeout",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
            }
            catch (HttpRequestException ex)
            {
                _logger.LogError(ex, "API connection test failed with HTTP error");

                System.Windows.MessageBox.Show(
                    $"API connection test failed!\n\n" +
                    $"URL: {ApiBaseUrl}\n" +
                    $"Error: {ex.Message}\n\n" +
                    $"Please check the URL format and your internet connection.",
                    "API Test Failed",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "API connection test failed with unexpected error");

                System.Windows.MessageBox.Show(
                    $"API connection test failed!\n\n" +
                    $"URL: {ApiBaseUrl}\n" +
                    $"Error: {ex.Message}\n\n" +
                    $"Please check the settings and try again.",
                    "API Test Failed",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        [RelayCommand]
        private async Task CheckForUpdates()
        {
            try
            {
                IsBusy = true;
                BusyMessage = "Checking for updates...";
                UpdateCheckStatus = "Checking for updates...";

                _logger.LogInformation("Manual update check initiated");

                var updateInfo = await _updateService.CheckForUpdatesAsync();
                if (updateInfo != null)
                {
                    // Update available
                    LatestVersion = updateInfo.Version.ToString();
                    UpdateAvailable = true;
                    UpdateCheckStatus = $"Update available! Version {updateInfo.Version} is ready to download.";
                    ReleaseNotes = updateInfo.ReleaseNotes ?? "No release notes available.";

                    _logger.LogInformation("Update available: Version {Version}", updateInfo.Version);
                    _logger.LogInformation("Release notes: {ReleaseNotes}", updateInfo.ReleaseNotes);

                    // Prompt user for installation
                    await PromptUserForUpdateInstallationAsync(updateInfo);
                }
                else
                {
                    // No updates available
                    var currentVer = _updateService.GetCurrentVersion();
                    LatestVersion = currentVer.ToString();
                    UpdateAvailable = false;
                    UpdateCheckStatus = $"You have the latest version ({currentVer}). No updates available.";
                    ReleaseNotes = string.Empty;

                    _logger.LogInformation("No updates available");
                }

                _logger.LogInformation("Update check completed");
            }
            catch (Exception ex)
            {
                UpdateCheckStatus = "Failed to check for updates. Please check your internet connection.";
                UpdateAvailable = false;
                ReleaseNotes = string.Empty;
                _logger.LogError(ex, "Failed to check for updates");
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void LoadVersionInformation()
        {
            try
            {
                var currentVer = _updateService.GetCurrentVersion();
                CurrentVersion = currentVer.ToString();
                LatestVersion = "Unknown - check for updates";
                UpdateAvailable = false;
                _logger.LogDebug("Loaded version information - Current: {Version}", currentVer);
            }
            catch (Exception ex)
            {
                CurrentVersion = "Unknown";
                LatestVersion = "Unknown";
                UpdateAvailable = false;
                _logger.LogError(ex, "Failed to load version information");
            }
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

        private async Task PromptUserForUpdateInstallationAsync(UpdateInfo updateInfo)
        {
            try
            {
                var message = $"Update Available!\n\nA new version ({updateInfo.Version}) is available for download.\n\nWould you like to install it now?\n\nRelease Notes:\n{updateInfo.ReleaseNotes}";
                var result = System.Windows.MessageBox.Show(
                    message,
                    "Update Available",
                    System.Windows.MessageBoxButton.YesNo,
                    System.Windows.MessageBoxImage.Information);

                if (result == System.Windows.MessageBoxResult.Yes)
                {
                    _logger.LogInformation("User accepted update installation for version {Version}", updateInfo.Version);

                    BusyMessage = "Installing update...";
                    IsBusy = true;

                    try
                    {
                        var installSuccess = await _updateService.DownloadAndInstallUpdateAsync(updateInfo, null);
                        if (installSuccess)
                        {
                            _logger.LogInformation("Update installation initiated successfully");
                            UpdateCheckStatus = "Update installation started. The application will restart to complete the update.";
                        }
                        else
                        {
                            _logger.LogWarning("Update installation failed");
                            UpdateCheckStatus = "Update installation failed. Please try again later.";
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error during update installation");
                        UpdateCheckStatus = "Update installation failed. Please check the logs for details.";
                    }
                    finally
                    {
                        IsBusy = false;
                    }
                }
                else
                {
                    _logger.LogInformation("User declined update installation for version {Version}", updateInfo.Version);
                    UpdateCheckStatus = $"Update available but not installed. Version {updateInfo.Version} can be installed later.";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error prompting user for update installation");
            }
        }
    }
}