using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;

namespace BulkEditor.Application.Services
{
    /// <summary>
    /// Manages automatic update checking and user interaction for updates
    /// </summary>
    public class UpdateManager : IDisposable
    {
        private readonly IUpdateService _updateService;
        private readonly IConfigurationService _configService;
        private readonly ILoggingService _logger;
        private readonly System.Timers.Timer _updateCheckTimer;
        private bool _disposed;
        private bool _updateInProgress;

        /// <summary>
        /// Event raised when an update is available
        /// </summary>
        public event EventHandler<UpdateAvailableEventArgs>? UpdateAvailable;

        /// <summary>
        /// Event raised when update progress changes
        /// </summary>
        public event EventHandler<UpdateProgressEventArgs>? UpdateProgressChanged;

        /// <summary>
        /// Event raised when update is completed
        /// </summary>
        public event EventHandler<UpdateCompletedEventArgs>? UpdateCompleted;

        public UpdateManager(
            IUpdateService updateService,
            IConfigurationService configService,
            ILoggingService logger)
        {
            _updateService = updateService ?? throw new ArgumentNullException(nameof(updateService));
            _configService = configService ?? throw new ArgumentNullException(nameof(configService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));

            // Initialize timer for periodic update checks (24 hours)
            _updateCheckTimer = new System.Timers.Timer(TimeSpan.FromHours(24).TotalMilliseconds);
            _updateCheckTimer.Elapsed += OnUpdateCheckTimerElapsed;
            _updateCheckTimer.AutoReset = true;
        }

        /// <summary>
        /// Start the update manager and begin periodic update checks
        /// </summary>
        public async Task StartAsync()
        {
            try
            {
                if (!_updateService.IsAutoUpdateEnabled())
                {
                    _logger.LogInformation("Auto-update is disabled, skipping update checks");
                    return;
                }

                _logger.LogInformation("Starting update manager");

                // Perform initial update check
                await CheckForUpdatesAsync();

                // Start the timer for periodic checks
                _updateCheckTimer.Start();

                _logger.LogInformation("Update manager started successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to start update manager");
            }
        }

        /// <summary>
        /// Stop the update manager
        /// </summary>
        public void Stop()
        {
            try
            {
                _updateCheckTimer?.Stop();
                _logger.LogInformation("Update manager stopped");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error stopping update manager");
            }
        }

        /// <summary>
        /// Manually check for updates
        /// </summary>
        public async Task<UpdateInfo> CheckForUpdatesAsync()
        {
            try
            {
                if (_updateInProgress)
                {
                    _logger.LogInformation("Update already in progress, skipping check");
                    return null;
                }

                _logger.LogInformation("Checking for application updates...");
                var updateInfo = await _updateService.CheckForUpdatesAsync();

                if (updateInfo != null)
                {
                    _logger.LogInformation("Update available: {Version}", updateInfo.Version);
                    UpdateAvailable?.Invoke(this, new UpdateAvailableEventArgs(updateInfo));
                }
                else
                {
                    _logger.LogInformation("No updates available");
                }

                return updateInfo;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking for updates");
                return null;
            }
        }

        /// <summary>
        /// Download and install an update
        /// </summary>
        public async Task<bool> InstallUpdateAsync(UpdateInfo updateInfo)
        {
            try
            {
                if (_updateInProgress)
                {
                    _logger.LogWarning("Update already in progress");
                    return false;
                }

                _updateInProgress = true;
                _logger.LogInformation("Starting update installation for version {Version}", updateInfo.Version);

                var progress = new Progress<UpdateProgress>(OnUpdateProgress);
                var success = await _updateService.DownloadAndInstallUpdateAsync(updateInfo, progress);

                UpdateCompleted?.Invoke(this, new UpdateCompletedEventArgs(success, updateInfo));

                return success;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to install update");
                UpdateCompleted?.Invoke(this, new UpdateCompletedEventArgs(false, updateInfo, ex.Message));
                return false;
            }
            finally
            {
                _updateInProgress = false;
            }
        }

        /// <summary>
        /// Enable or disable automatic update checking
        /// </summary>
        public void SetAutoUpdateEnabled(bool enabled)
        {
            _updateService.SetAutoUpdateEnabled(enabled);

            if (enabled)
            {
                _updateCheckTimer.Start();
            }
            else
            {
                _updateCheckTimer.Stop();
            }
        }

        /// <summary>
        /// Get the current application version
        /// </summary>
        public Version GetCurrentVersion()
        {
            return _updateService.GetCurrentVersion();
        }

        private async void OnUpdateCheckTimerElapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                await CheckForUpdatesAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during scheduled update check");
            }
        }

        private void OnUpdateProgress(UpdateProgress progress)
        {
            UpdateProgressChanged?.Invoke(this, new UpdateProgressEventArgs(progress));
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _updateCheckTimer?.Stop();
                _updateCheckTimer?.Dispose();
                _disposed = true;
            }
        }
    }

    /// <summary>
    /// Event arguments for update available notification
    /// </summary>
    public class UpdateAvailableEventArgs : EventArgs
    {
        public UpdateInfo UpdateInfo { get; }

        public UpdateAvailableEventArgs(UpdateInfo updateInfo)
        {
            UpdateInfo = updateInfo;
        }
    }

    /// <summary>
    /// Event arguments for update progress notification
    /// </summary>
    public class UpdateProgressEventArgs : EventArgs
    {
        public UpdateProgress Progress { get; }

        public UpdateProgressEventArgs(UpdateProgress progress)
        {
            Progress = progress;
        }
    }

    /// <summary>
    /// Event arguments for update completion notification
    /// </summary>
    public class UpdateCompletedEventArgs : EventArgs
    {
        public bool Success { get; }
        public UpdateInfo UpdateInfo { get; }
        public string ErrorMessage { get; }

        public UpdateCompletedEventArgs(bool success, UpdateInfo updateInfo, string errorMessage = null)
        {
            Success = success;
            UpdateInfo = updateInfo;
            ErrorMessage = errorMessage;
        }
    }
}