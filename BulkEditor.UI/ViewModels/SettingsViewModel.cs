using BulkEditor.Core.Configuration;
using BulkEditor.Core.Interfaces;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BulkEditor.UI.ViewModels
{
    /// <summary>
    /// ViewModel for the Settings window
    /// </summary>
    public partial class SettingsViewModel : ViewModelBase
    {
        private readonly ILoggingService _logger;
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

        // Replacement Rules Collections
        public ObservableCollection<HyperlinkReplacementRule> HyperlinkRules { get; set; } = new();
        public ObservableCollection<TextReplacementRule> TextRules { get; set; } = new();

        public SettingsViewModel(AppSettings appSettings, ILoggingService logger)
        {
            _originalSettings = appSettings ?? throw new ArgumentNullException(nameof(appSettings));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));

            // Create a copy of the settings to work with
            _currentSettings = CloneSettings(appSettings);

            Title = "Settings";
            LoadSettingsIntoProperties();
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

                // Update current settings from UI properties
                UpdateSettingsFromProperties();

                // Validate settings
                if (!ValidateSettings())
                {
                    IsBusy = false;
                    return;
                }

                // Copy current settings back to original
                CopySettingsTo(_currentSettings, _originalSettings);

                // Save settings to file
                await SaveSettingsToFileAsync();

                _logger.LogInformation("Settings saved successfully");

                IsBusy = false;
                RequestClose?.Invoke(this, true);
            }
            catch (Exception ex)
            {
                IsBusy = false;
                _logger.LogError(ex, "Error saving settings");
                // The error will be shown via the notification service in the parent window
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

            return true;
        }

        private async Task SaveSettingsToFileAsync()
        {
            try
            {
                var settingsPath = Path.Combine(Environment.CurrentDirectory, "appsettings.json");
                var json = System.Text.Json.JsonSerializer.Serialize(_currentSettings, new System.Text.Json.JsonSerializerOptions
                {
                    WriteIndented = true
                });

                await File.WriteAllTextAsync(settingsPath, json);
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
    }
}