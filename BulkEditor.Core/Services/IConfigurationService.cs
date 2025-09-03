using BulkEditor.Core.Configuration;
using System.Threading.Tasks;

namespace BulkEditor.Core.Services
{
    /// <summary>
    /// Service for managing application configuration and settings
    /// </summary>
    public interface IConfigurationService
    {
        /// <summary>
        /// Load application settings from the configuration file
        /// </summary>
        Task<AppSettings> LoadSettingsAsync();

        /// <summary>
        /// Save application settings to the configuration file
        /// </summary>
        Task SaveSettingsAsync(AppSettings settings);

        /// <summary>
        /// Get the path to the application data directory
        /// </summary>
        string GetAppDataPath();

        /// <summary>
        /// Get the path to the settings file
        /// </summary>
        string GetSettingsFilePath();

        /// <summary>
        /// Initialize the application data directory structure
        /// </summary>
        Task InitializeAsync();

        /// <summary>
        /// Create a backup of current settings
        /// </summary>
        Task BackupSettingsAsync(string backupSuffix = null);

        /// <summary>
        /// Restore settings from a backup
        /// </summary>
        Task RestoreSettingsAsync(string backupFilePath);

        /// <summary>
        /// Check if this is the first run of the application
        /// </summary>
        bool IsFirstRun();

        /// <summary>
        /// Migrate settings from an older version or location
        /// </summary>
        Task MigrateSettingsAsync();
    }
}