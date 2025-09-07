using System;
using System.Threading.Tasks;

namespace BulkEditor.Core.Services
{
    /// <summary>
    /// Service for managing application updates
    /// </summary>
    public interface IUpdateService
    {
        /// <summary>
        /// Check if an update is available on GitHub
        /// </summary>
        Task<UpdateInfo> CheckForUpdatesAsync();

        /// <summary>
        /// Download and install the latest update
        /// </summary>
        Task<bool> DownloadAndInstallUpdateAsync(UpdateInfo updateInfo, IProgress<UpdateProgress>? progress = null);

        /// <summary>
        /// Get the current application version
        /// </summary>
        Version GetCurrentVersion();

        /// <summary>
        /// Enable or disable automatic update checking
        /// </summary>
        void SetAutoUpdateEnabled(bool enabled);

        /// <summary>
        /// Check if auto-update is enabled
        /// </summary>
        bool IsAutoUpdateEnabled();

        /// <summary>
        /// Prepare the application for update (backup settings, close files, etc.)
        /// </summary>
        Task PrepareForUpdateAsync();

        /// <summary>
        /// Rollback to previous version if update fails
        /// </summary>
        Task<bool> RollbackUpdateAsync();
    }

    /// <summary>
    /// Information about an available update
    /// </summary>
    public class UpdateInfo
    {
        public Version Version { get; set; } = new Version();
        public string DownloadUrl { get; set; } = string.Empty;
        public string ReleaseNotes { get; set; } = string.Empty;
        public DateTime ReleaseDate { get; set; }
        public long FileSize { get; set; }
        public string FileName { get; set; } = string.Empty;
        public bool IsPrerelease { get; set; }
        public bool IsSecurityUpdate { get; set; }
    }

    /// <summary>
    /// Progress information for update operations
    /// </summary>
    public class UpdateProgress
    {
        public UpdateProgressType Type { get; set; }
        public string Message { get; set; } = string.Empty;
        public int PercentComplete { get; set; }
        public long BytesDownloaded { get; set; }
        public long TotalBytes { get; set; }
    }

    /// <summary>
    /// Types of update progress
    /// </summary>
    public enum UpdateProgressType
    {
        Checking,
        Downloading,
        Installing,
        BackingUp,
        Finalizing,
        Complete,
        Error
    }
}