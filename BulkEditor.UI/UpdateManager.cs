using BulkEditor.Core.Configuration;
using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using BulkEditor.UI.Services;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace BulkEditor.UI
{
    /// <summary>
    /// Manages the application update process, checking for updates at startup and periodically.
    /// </summary>
    public class UpdateManager : IDisposable
    {
        private readonly IUpdateService _updateService;
        private readonly ILoggingService _logger;
        private readonly INotificationService _notificationService;
        private readonly AppSettings _appSettings;
        private Timer _updateTimer;

        public UpdateManager(
            IUpdateService updateService,
            ILoggingService logger,
            INotificationService notificationService,
            AppSettings appSettings)
        {
            _updateService = updateService ?? throw new ArgumentNullException(nameof(updateService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _notificationService = notificationService ?? throw new ArgumentNullException(nameof(notificationService));
            _appSettings = appSettings ?? throw new ArgumentNullException(nameof(appSettings));
        }

        /// <summary>
        /// Starts the update manager, performing an initial check and setting up periodic checks.
        /// </summary>
        public async Task StartAsync()
        {
            _logger.LogInformation("Update Manager starting...");

            // Perform initial check on startup
            await CheckForUpdatesAsync(true);

            // Set up periodic checks
            var checkInterval = TimeSpan.FromHours(_appSettings.Update.CheckIntervalHours);
            if (checkInterval > TimeSpan.Zero)
            {
                _updateTimer = new Timer(async _ => await CheckForUpdatesAsync(false), null, checkInterval, checkInterval);
                _logger.LogInformation("Scheduled periodic update checks every {Hours} hours.", _appSettings.Update.CheckIntervalHours);
            }
        }

        /// <summary>
        /// Stops the periodic update checks.
        /// </summary>
        public void Stop()
        {
            _logger.LogInformation("Update Manager stopping...");
            _updateTimer?.Change(Timeout.Infinite, 0);
        }

        private async Task CheckForUpdatesAsync(bool isStartup)
        {
            if (!_appSettings.Update.AutoUpdateEnabled)
            {
                if (isStartup) _logger.LogInformation("Automatic update check is disabled. Skipping startup check.");
                return;
            }

            try
            {
                _logger.LogInformation("Checking for application updates...");
                var updateInfo = await _updateService.CheckForUpdatesAsync();

                if (updateInfo != null)
                {
                    _logger.LogInformation("Update found: Version {Version}", updateInfo.Version);

                    var message = $"A new version ({updateInfo.Version}) is available. Do you want to download and install it now?";

                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        _notificationService.ShowActionableNotification("Update Available", message, "Update Now", async () =>
                        {
                            await _updateService.DownloadAndInstallUpdateAsync(updateInfo);
                        });
                    });
                }
                else
                {
                    _logger.LogInformation("Application is up-to-date.");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An error occurred while checking for updates.");
            }
        }

        public void Dispose()
        {
            _updateTimer?.Dispose();
        }
    }
}
