using BulkEditor.Core.Configuration;
using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Service for managing application configuration and settings in AppData
    /// </summary>
    public class ConfigurationService : IConfigurationService
    {
        private readonly ILoggingService _logger;
        private readonly string _appDataPath;
        private readonly string _settingsFileName = "settings.json";
        private readonly string _backupFolderName = "SettingsBackups";

        public ConfigurationService(ILoggingService logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "BulkEditor");
        }

        public string GetAppDataPath() => _appDataPath;

        public string GetSettingsFilePath() => Path.Combine(_appDataPath, _settingsFileName);

        public bool IsFirstRun()
        {
            return !File.Exists(GetSettingsFilePath());
        }

        public async Task InitializeAsync()
        {
            try
            {
                // Create main app data directory
                if (!Directory.Exists(_appDataPath))
                {
                    Directory.CreateDirectory(_appDataPath);
                    _logger.LogInformation("Created app data directory: {Path}", _appDataPath);
                }

                // Create subdirectories
                var subdirectories = new[]
                {
                    "Logs",
                    "Backups",
                    "Cache",
                    "Temp",
                    _backupFolderName
                };

                foreach (var subdir in subdirectories)
                {
                    var path = Path.Combine(_appDataPath, subdir);
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                        _logger.LogInformation("Created subdirectory: {Path}", path);
                    }
                }

                // If this is first run, create default settings
                if (IsFirstRun())
                {
                    await CreateDefaultSettingsAsync();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to initialize application data directories");
                throw;
            }
        }

        public async Task<AppSettings> LoadSettingsAsync()
        {
            try
            {
                var settingsPath = GetSettingsFilePath();

                if (!File.Exists(settingsPath))
                {
                    _logger.LogInformation("Settings file not found, creating default settings");
                    return await CreateDefaultSettingsAsync();
                }

                var json = await File.ReadAllTextAsync(settingsPath);
                var settings = JsonSerializer.Deserialize<AppSettings>(json, GetJsonOptions());

                if (settings == null)
                {
                    _logger.LogWarning("Failed to deserialize settings, using defaults");
                    return await CreateDefaultSettingsAsync();
                }

                // Enhanced logging: Log critical settings after loading
                _logger.LogInformation("Settings loaded from file - AutoReplaceTitles: {AutoReplace}, EnableHyperlinkReplacement: {HyperlinkReplace}, HyperlinkRuleCount: {RuleCount}",
                    settings.Validation.AutoReplaceTitles, settings.Replacement.EnableHyperlinkReplacement, settings.Replacement.HyperlinkRules?.Count ?? 0);

                // Update paths to use AppData locations
                UpdatePathsToAppData(settings);

                _logger.LogInformation("Settings loaded successfully from: {Path}", settingsPath);
                return settings;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to load settings from: {Path}", GetSettingsFilePath());
                return await CreateDefaultSettingsAsync();
            }
        }

        public async Task SaveSettingsAsync(AppSettings settings)
        {
            try
            {
                // Enhanced logging: Log critical settings before saving
                _logger.LogInformation("Saving settings to file - AutoReplaceTitles: {AutoReplace}, EnableHyperlinkReplacement: {HyperlinkReplace}, HyperlinkRuleCount: {RuleCount}",
                    settings.Validation.AutoReplaceTitles, settings.Replacement.EnableHyperlinkReplacement, settings.Replacement.HyperlinkRules?.Count ?? 0);

                // Create backup before saving
                await BackupSettingsAsync("pre-save");

                var settingsPath = GetSettingsFilePath();
                var json = JsonSerializer.Serialize(settings, GetJsonOptions());

                await File.WriteAllTextAsync(settingsPath, json);
                _logger.LogInformation("Settings saved successfully to: {Path}", settingsPath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to save settings to: {Path}", GetSettingsFilePath());
                throw;
            }
        }

        public async Task BackupSettingsAsync(string? backupSuffix = null)
        {
            try
            {
                var settingsPath = GetSettingsFilePath();
                if (!File.Exists(settingsPath))
                {
                    _logger.LogInformation("No settings file to backup");
                    return;
                }

                var backupDir = Path.Combine(_appDataPath, _backupFolderName);
                Directory.CreateDirectory(backupDir);

                var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                var suffix = string.IsNullOrEmpty(backupSuffix) ? "" : $"_{backupSuffix}";
                var backupFileName = $"settings_backup_{timestamp}{suffix}.json";
                var backupPath = Path.Combine(backupDir, backupFileName);

                await FileExtensions.CopyAsync(settingsPath, backupPath);
                _logger.LogInformation("Settings backed up to: {Path}", backupPath);

                // Clean up old backups (keep last 10)
                CleanupOldBackups();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to backup settings");
                // Don't throw - backup failure shouldn't prevent saving
            }
        }

        public async Task RestoreSettingsAsync(string backupFilePath)
        {
            try
            {
                if (!File.Exists(backupFilePath))
                {
                    throw new FileNotFoundException($"Backup file not found: {backupFilePath}");
                }

                // Backup current settings before restore
                await BackupSettingsAsync("pre-restore");

                var settingsPath = GetSettingsFilePath();
                await FileExtensions.CopyAsync(backupFilePath, settingsPath, true);

                _logger.LogInformation("Settings restored from backup: {BackupPath}", backupFilePath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to restore settings from backup: {BackupPath}", backupFilePath);
                throw;
            }
        }

        public async Task MigrateSettingsAsync()
        {
            try
            {
                // Check for settings in old location (current directory)
                var oldSettingsPath = Path.Combine(Environment.CurrentDirectory, "appsettings.json");

                if (File.Exists(oldSettingsPath) && IsFirstRun())
                {
                    _logger.LogInformation("Migrating settings from old location: {OldPath}", oldSettingsPath);

                    var oldJson = await File.ReadAllTextAsync(oldSettingsPath);
                    var oldSettings = JsonSerializer.Deserialize<AppSettings>(oldJson, GetJsonOptions());

                    if (oldSettings != null)
                    {
                        // Update paths to new AppData locations
                        UpdatePathsToAppData(oldSettings);

                        // Save to new location
                        await SaveSettingsAsync(oldSettings);

                        // Create backup of old file in case we need to rollback
                        var migrationBackupPath = Path.Combine(_appDataPath, _backupFolderName, $"migration_backup_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.json");
                        await FileExtensions.CopyAsync(oldSettingsPath, migrationBackupPath);

                        _logger.LogInformation("Settings migration completed successfully");
                        return;
                    }
                }

                _logger.LogInformation("No settings migration needed");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to migrate settings");
                // Don't throw - migration failure shouldn't prevent app startup
            }
        }

        private async Task<AppSettings> CreateDefaultSettingsAsync()
        {
            var defaultSettings = new AppSettings
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
                    BackupDirectory = Path.Combine(_appDataPath, "Backups"),
                    CreateTimestampedBackups = true,
                    MaxBackupAge = 30,
                    CompressBackups = false,
                    AutoCleanupOldBackups = true
                },
                Logging = new LoggingSettings
                {
                    LogLevel = "Information",
                    LogDirectory = Path.Combine(_appDataPath, "Logs"),
                    EnableFileLogging = true,
                    EnableConsoleLogging = true,
                    MaxLogFileSizeMB = 10,
                    MaxLogFiles = 5,
                    LogFilePattern = "bulkeditor-{Date}.log"
                },
                UI = new UiSettings
                {
                    Theme = "Light",
                    Language = "en-US",
                    ShowProgressDetails = true,
                    MinimizeToSystemTray = false,
                    ConfirmBeforeProcessing = true,
                    AutoSaveSettings = true,
                    Window = new WindowSettings
                    {
                        Width = 1200,
                        Height = 800,
                        Left = 100,
                        Top = 100,
                        Maximized = false
                    }
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

            await SaveSettingsAsync(defaultSettings);
            return defaultSettings;
        }

        private void UpdatePathsToAppData(AppSettings settings)
        {
            // Update backup directory to AppData if it's a relative path
            if (!Path.IsPathRooted(settings.Backup.BackupDirectory))
            {
                settings.Backup.BackupDirectory = Path.Combine(_appDataPath, "Backups");
            }

            // Update log directory to AppData if it's a relative path
            if (!Path.IsPathRooted(settings.Logging.LogDirectory))
            {
                settings.Logging.LogDirectory = Path.Combine(_appDataPath, "Logs");
            }
        }

        private static JsonSerializerOptions GetJsonOptions()
        {
            return new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
            };
        }

        private void CleanupOldBackups()
        {
            try
            {
                var backupDir = Path.Combine(_appDataPath, _backupFolderName);
                if (!Directory.Exists(backupDir))
                    return;

                var backupFiles = Directory.GetFiles(backupDir, "settings_backup_*.json");
                if (backupFiles.Length <= 10)
                    return;

                // Sort by creation time and keep only the 10 most recent
                Array.Sort(backupFiles, (x, y) => File.GetCreationTime(y).CompareTo(File.GetCreationTime(x)));

                for (int i = 10; i < backupFiles.Length; i++)
                {
                    File.Delete(backupFiles[i]);
                    _logger.LogInformation("Deleted old backup: {Path}", backupFiles[i]);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to cleanup old backup files");
            }
        }
    }

    /// <summary>
    /// Extension method for File.CopyAsync which doesn't exist in .NET
    /// </summary>
    public static class FileExtensions
    {
        public static async Task CopyAsync(string sourceFilePath, string destinationFilePath, bool overwrite = false)
        {
            using var sourceStream = new FileStream(sourceFilePath, FileMode.Open, FileAccess.Read);
            using var destinationStream = new FileStream(destinationFilePath,
                overwrite ? FileMode.Create : FileMode.CreateNew, FileAccess.Write);

            await sourceStream.CopyToAsync(destinationStream);
        }
    }
}