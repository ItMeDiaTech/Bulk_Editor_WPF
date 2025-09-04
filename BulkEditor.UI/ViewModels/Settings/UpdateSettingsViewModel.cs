using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Threading.Tasks;

namespace BulkEditor.UI.ViewModels.Settings
{
    public partial class UpdateSettingsViewModel : ObservableObject
    {
        private readonly ILoggingService? _logger;
        private readonly IUpdateService? _updateService;

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

        [ObservableProperty]
        private bool _isBusy;

        [ObservableProperty]
        private string _busyMessage = string.Empty;

        public UpdateSettingsViewModel(ILoggingService? logger = null, IUpdateService? updateService = null)
        {
            _logger = logger;
            _updateService = updateService;
            LoadVersionInformation();
        }

        [RelayCommand]
        private async Task CheckForUpdates()
        {
            try
            {
                IsBusy = true;
                BusyMessage = "Checking for updates...";
                UpdateCheckStatus = "Checking for updates...";

                _logger?.LogInformation("Manual update check initiated");

                if (_updateService == null)
                {
                    _logger?.LogWarning("Update service is not available");
                    UpdateCheckStatus = "Update service is not available.";
                    return;
                }

                var updateInfo = await _updateService.CheckForUpdatesAsync();
                if (updateInfo != null)
                {
                    // Update available
                    LatestVersion = updateInfo.Version.ToString();
                    UpdateAvailable = true;
                    UpdateCheckStatus = $"Update available! Version {updateInfo.Version} is ready to download.";
                    ReleaseNotes = updateInfo.ReleaseNotes ?? "No release notes available.";

                    _logger?.LogInformation("Update available: Version {Version}", updateInfo.Version);
                    _logger?.LogInformation("Release notes: {ReleaseNotes}", updateInfo.ReleaseNotes);

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

                    _logger?.LogInformation("No updates available");
                }

                _logger?.LogInformation("Update check completed");
            }
            catch (Exception ex)
            {
                UpdateCheckStatus = "Failed to check for updates. Please check your internet connection.";
                UpdateAvailable = false;
                ReleaseNotes = string.Empty;
                _logger?.LogError(ex, "Failed to check for updates");
            }
            finally
            {
                IsBusy = false;
                BusyMessage = string.Empty;
            }
        }

        private void LoadVersionInformation()
        {
            try
            {
                if (_updateService == null)
                {
                    CurrentVersion = "Unknown";
                    LatestVersion = "Unknown";
                    UpdateAvailable = false;
                    return;
                }

                var currentVer = _updateService.GetCurrentVersion();
                CurrentVersion = currentVer.ToString();
                LatestVersion = "Unknown - check for updates";
                UpdateAvailable = false;
                _logger?.LogDebug("Loaded version information - Current: {Version}", currentVer);
            }
            catch (Exception ex)
            {
                CurrentVersion = "Unknown";
                LatestVersion = "Unknown";
                UpdateAvailable = false;
                _logger?.LogError(ex, "Failed to load version information");
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
                    _logger?.LogInformation("User accepted update installation for version {Version}", updateInfo.Version);

                    BusyMessage = "Installing update...";
                    IsBusy = true;

                    try
                    {
                        if (_updateService == null)
                        {
                            _logger?.LogWarning("Update service is not available for installation");
                            UpdateCheckStatus = "Update service is not available for installation.";
                            return;
                        }

                        var installSuccess = await _updateService.DownloadAndInstallUpdateAsync(updateInfo, null);
                        if (installSuccess)
                        {
                            _logger?.LogInformation("Update installation initiated successfully");
                            UpdateCheckStatus = "Update installation started. The application will restart to complete the update.";
                        }
                        else
                        {
                            _logger?.LogWarning("Update installation failed");
                            UpdateCheckStatus = "Update installation failed. Please try again later.";
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogError(ex, "Error during update installation");
                        UpdateCheckStatus = "Update installation failed. Please check the logs for details.";
                    }
                    finally
                    {
                        IsBusy = false;
                        BusyMessage = string.Empty;
                    }
                }
                else
                {
                    _logger?.LogInformation("User declined update installation for version {Version}", updateInfo.Version);
                    UpdateCheckStatus = $"Update available but not installed. Version {updateInfo.Version} can be installed later.";
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error prompting user for update installation");
            }
        }
    }
}