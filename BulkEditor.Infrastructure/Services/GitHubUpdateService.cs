using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Service for checking and downloading updates from GitHub releases
    /// </summary>
    public class GitHubUpdateService : IUpdateService
    {
        private readonly HttpClient _httpClient;
        private readonly ILoggingService _logger;
        private readonly IConfigurationService _configService;
        private readonly string _githubOwner;
        private readonly string _githubRepo;
        private readonly string _appDataPath;
        private bool _autoUpdateEnabled;

        /// <summary>
        /// Event raised when the application needs to shut down for update installation
        /// </summary>
        public event EventHandler? UpdateRequiresRestart;

        public GitHubUpdateService(
            HttpClient httpClient,
            ILoggingService logger,
            IConfigurationService configService,
            string githubOwner = "ItMeDiaTech",
            string githubRepo = "Bulk_Editor_WPF")
        {
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _configService = configService ?? throw new ArgumentNullException(nameof(configService));
            _githubOwner = githubOwner;
            _githubRepo = githubRepo;
            _appDataPath = _configService.GetAppDataPath();
            _autoUpdateEnabled = true; // Default to enabled
        }

        public Version GetCurrentVersion()
        {
            try
            {
                // Get the UI assembly version since that's the main application
                var entryAssembly = Assembly.GetEntryAssembly();
                if (entryAssembly != null)
                {
                    var version = entryAssembly.GetName().Version;
                    if (version != null)
                    {
                        _logger.LogDebug("Current version from entry assembly: {Version}", version);
                        return version;
                    }
                }

                // Fallback to executing assembly
                var assembly = Assembly.GetExecutingAssembly();
                var fallbackVersion = assembly.GetName().Version;
                _logger.LogDebug("Current version from executing assembly: {Version}", fallbackVersion);
                return fallbackVersion ?? new Version(1, 0, 10, 0); // Match current release
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to get current version");
                return new Version(1, 0, 10, 0); // Match current release
            }
        }

        public bool IsAutoUpdateEnabled() => _autoUpdateEnabled;

        public void SetAutoUpdateEnabled(bool enabled)
        {
            _autoUpdateEnabled = enabled;
            _logger.LogInformation("Auto-update {Status}", enabled ? "enabled" : "disabled");
        }

        public async Task<UpdateInfo> CheckForUpdatesAsync()
        {
            try
            {
                _logger.LogInformation("Checking for updates from GitHub...");

                var apiUrl = $"https://api.github.com/repos/{_githubOwner}/{_githubRepo}/releases/latest";
                var response = await _httpClient.GetAsync(apiUrl);

                if (!response.IsSuccessStatusCode)
                {
                    _logger.LogWarning("Failed to check for updates. Status: {StatusCode}", response.StatusCode);
                    return null;
                }

                var jsonContent = await response.Content.ReadAsStringAsync();
                var release = JsonSerializer.Deserialize<GitHubRelease>(jsonContent, new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower
                });

                if (release == null)
                {
                    _logger.LogWarning("Failed to parse GitHub release information");
                    return null;
                }

                // Parse version from tag name (assuming format "v1.2.3" or "1.2.3")
                var versionString = release.TagName.TrimStart('v');
                if (!Version.TryParse(versionString, out var releaseVersion))
                {
                    _logger.LogWarning("Could not parse version from tag: {TagName}", release.TagName);
                    return null;
                }

                var currentVersion = GetCurrentVersion();
                if (releaseVersion <= currentVersion)
                {
                    _logger.LogInformation("No updates available. Current: {Current}, Latest: {Latest}",
                        currentVersion, releaseVersion);
                    return null;
                }

                // Find installer asset
                GitHubAsset installerAsset = null;
                foreach (var asset in release.Assets)
                {
                    if (asset.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) ||
                        asset.Name.EndsWith(".msi", StringComparison.OrdinalIgnoreCase))
                    {
                        installerAsset = asset;
                        break;
                    }
                }

                if (installerAsset == null)
                {
                    _logger.LogWarning("No installer found in release assets");
                    return null;
                }

                var updateInfo = new UpdateInfo
                {
                    Version = releaseVersion,
                    DownloadUrl = installerAsset.BrowserDownloadUrl,
                    ReleaseNotes = release.Body,
                    ReleaseDate = release.PublishedAt,
                    FileSize = installerAsset.Size,
                    FileName = installerAsset.Name,
                    IsPrerelease = release.Prerelease,
                    IsSecurityUpdate = release.Body?.Contains("security", StringComparison.OrdinalIgnoreCase) == true
                };

                _logger.LogInformation("Update available: {Version}", releaseVersion);
                return updateInfo;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking for updates");
                return null;
            }
        }

        public async Task<bool> DownloadAndInstallUpdateAsync(UpdateInfo updateInfo, IProgress<UpdateProgress> progress = null)
        {
            try
            {
                if (updateInfo == null)
                {
                    throw new ArgumentNullException(nameof(updateInfo));
                }

                _logger.LogInformation("Starting update download and installation for version {Version}", updateInfo.Version);

                // Prepare for update
                progress?.Report(new UpdateProgress
                {
                    Type = UpdateProgressType.BackingUp,
                    Message = "Preparing for update...",
                    PercentComplete = 0
                });

                await PrepareForUpdateAsync();

                // Download the installer
                progress?.Report(new UpdateProgress
                {
                    Type = UpdateProgressType.Downloading,
                    Message = "Downloading update...",
                    PercentComplete = 10
                });

                var downloadPath = await DownloadUpdateAsync(updateInfo, progress);
                if (string.IsNullOrEmpty(downloadPath))
                {
                    return false;
                }

                // Install the update
                progress?.Report(new UpdateProgress
                {
                    Type = UpdateProgressType.Installing,
                    Message = "Installing update...",
                    PercentComplete = 80
                });

                var installSuccess = await InstallUpdateAsync(downloadPath, progress);

                progress?.Report(new UpdateProgress
                {
                    Type = UpdateProgressType.Complete,
                    Message = installSuccess ? "Update completed successfully" : "Update failed",
                    PercentComplete = 100
                });

                return installSuccess;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to download and install update");
                progress?.Report(new UpdateProgress
                {
                    Type = UpdateProgressType.Error,
                    Message = $"Update failed: {ex.Message}",
                    PercentComplete = 0
                });
                return false;
            }
        }

        public async Task PrepareForUpdateAsync()
        {
            try
            {
                _logger.LogInformation("Preparing for update...");

                // Create backup of settings
                await _configService.BackupSettingsAsync("pre-update");

                // Create backup of current installation
                await CreateInstallationBackupAsync();

                _logger.LogInformation("Update preparation completed");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to prepare for update");
                throw;
            }
        }

        public async Task<bool> RollbackUpdateAsync()
        {
            try
            {
                _logger.LogInformation("Starting update rollback...");

                var backupDir = Path.Combine(_appDataPath, "UpdateBackups");
                var latestBackup = GetLatestBackup(backupDir);

                if (string.IsNullOrEmpty(latestBackup))
                {
                    _logger.LogWarning("No backup found for rollback");
                    return false;
                }

                // Restore from backup would require more complex logic
                // For now, just restore settings
                var settingsBackup = Path.Combine(_appDataPath, "SettingsBackups");
                var latestSettingsBackup = GetLatestBackup(settingsBackup);

                if (!string.IsNullOrEmpty(latestSettingsBackup))
                {
                    await _configService.RestoreSettingsAsync(latestSettingsBackup);
                    _logger.LogInformation("Settings restored from backup");
                }

                _logger.LogInformation("Rollback completed");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to rollback update");
                return false;
            }
        }

        private async Task<string> DownloadUpdateAsync(UpdateInfo updateInfo, IProgress<UpdateProgress> progress)
        {
            try
            {
                var downloadsDir = Path.Combine(_appDataPath, "Downloads");
                Directory.CreateDirectory(downloadsDir);

                var downloadPath = Path.Combine(downloadsDir, updateInfo.FileName);

                using var response = await _httpClient.GetAsync(updateInfo.DownloadUrl, HttpCompletionOption.ResponseHeadersRead);
                response.EnsureSuccessStatusCode();

                using var contentStream = await response.Content.ReadAsStreamAsync();
                using var fileStream = new FileStream(downloadPath, FileMode.Create, FileAccess.Write, FileShare.None);

                var buffer = new byte[8192];
                long totalBytesRead = 0;
                int bytesRead;

                while ((bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                {
                    await fileStream.WriteAsync(buffer, 0, bytesRead);
                    totalBytesRead += bytesRead;

                    if (updateInfo.FileSize > 0)
                    {
                        var percentComplete = (int)((totalBytesRead * 70) / updateInfo.FileSize) + 10; // 10-80%
                        progress?.Report(new UpdateProgress
                        {
                            Type = UpdateProgressType.Downloading,
                            Message = "Downloading update...",
                            PercentComplete = Math.Min(percentComplete, 80),
                            BytesDownloaded = totalBytesRead,
                            TotalBytes = updateInfo.FileSize
                        });
                    }
                }

                _logger.LogInformation("Update downloaded to: {Path}", downloadPath);
                return downloadPath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to download update");
                return string.Empty;
            }
        }

        private async Task<bool> InstallUpdateAsync(string installerPath, IProgress<UpdateProgress> progress)
        {
            try
            {
                // Create a batch script to handle the installation after this process exits
                var batchPath = Path.Combine(_appDataPath, "update_installer.bat");
                var currentExePath = Process.GetCurrentProcess().MainModule?.FileName;

                var batchContent = $@"
@echo off
echo Installing BulkEditor update...
timeout /t 3 /nobreak > nul

""{installerPath}"" /S /D=""{Path.GetDirectoryName(currentExePath)}""

echo Update installation completed.
del ""{installerPath}""
del ""%~f0""
";

                await File.WriteAllTextAsync(batchPath, batchContent);

                // Start the batch file and exit the current application
                var startInfo = new ProcessStartInfo
                {
                    FileName = batchPath,
                    UseShellExecute = true,
                    CreateNoWindow = false
                };

                Process.Start(startInfo);

                _logger.LogInformation("Update installer started, application will now exit");

                // Signal the application to shut down
                UpdateRequiresRestart?.Invoke(this, EventArgs.Empty);

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to install update");
                return false;
            }
        }

        private async Task CreateInstallationBackupAsync()
        {
            try
            {
                var backupDir = Path.Combine(_appDataPath, "UpdateBackups", $"backup_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}");
                Directory.CreateDirectory(backupDir);

                var currentExeDir = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule?.FileName);
                if (!string.IsNullOrEmpty(currentExeDir))
                {
                    // For a full backup, we'd copy the entire installation directory
                    // For now, just log that we would do this
                    _logger.LogInformation("Installation backup directory created: {BackupDir}", backupDir);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to create installation backup");
                throw;
            }
        }

        private string GetLatestBackup(string backupDirectory)
        {
            try
            {
                if (!Directory.Exists(backupDirectory))
                    return string.Empty;

                var backupFiles = Directory.GetFiles(backupDirectory, "*backup*.json");
                if (backupFiles.Length == 0)
                    return string.Empty;

                Array.Sort(backupFiles, (x, y) => File.GetCreationTime(y).CompareTo(File.GetCreationTime(x)));
                return backupFiles[0];
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to find latest backup");
                return string.Empty;
            }
        }
    }

    // GitHub API response models
    internal class GitHubRelease
    {
        public string TagName { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Body { get; set; } = string.Empty;
        public bool Prerelease { get; set; }
        public DateTime PublishedAt { get; set; }
        public GitHubAsset[] Assets { get; set; } = Array.Empty<GitHubAsset>();
    }

    internal class GitHubAsset
    {
        public string Name { get; set; } = string.Empty;
        public string BrowserDownloadUrl { get; set; } = string.Empty;
        public long Size { get; set; }
        public string ContentType { get; set; } = string.Empty;
    }
}