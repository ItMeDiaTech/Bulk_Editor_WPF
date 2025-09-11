using BulkEditor.Core.Configuration;
using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using BulkEditor.UI.ViewModels;
using Moq;
using System.Threading.Tasks;
using Xunit;

namespace BulkEditor.Tests.UI.ViewModels
{
    /// <summary>
    /// Tests for SettingsViewModel functionality
    /// </summary>
    public class SettingsViewModelTests
    {
        private readonly Mock<ILoggingService> _mockLogger;
        private readonly Mock<IConfigurationService> _mockConfigurationService;
        private readonly Mock<IUpdateService> _mockUpdateService;
        private readonly Mock<IHttpService> _mockHttpService;
        private readonly Mock<IThemeService> _mockThemeService;
        private readonly AppSettings _appSettings;
        private readonly SettingsViewModel _viewModel;

        public SettingsViewModelTests()
        {
            _mockLogger = new Mock<ILoggingService>();
            _mockConfigurationService = new Mock<IConfigurationService>();
            _mockUpdateService = new Mock<IUpdateService>();
            _mockHttpService = new Mock<IHttpService>();
            _mockThemeService = new Mock<IThemeService>();
            
            // Setup theme service mock
            _mockThemeService.Setup(x => x.AvailableThemes).Returns(new[] { "Light", "Purple", "Pink", "Light Blue", "Green", "Dark", "Auto" });
            _mockThemeService.Setup(x => x.CurrentTheme).Returns("Light");
            
            _appSettings = CreateTestAppSettings();
            _viewModel = new SettingsViewModel(_appSettings, _mockLogger.Object, _mockConfigurationService.Object, _mockUpdateService.Object, _mockHttpService.Object, _mockThemeService.Object);
        }

        [Fact]
        public void Constructor_ShouldInitializePropertiesFromAppSettings()
        {
            // Assert
            Assert.Equal(_appSettings.Processing.MaxConcurrentDocuments, _viewModel.ProcessingSettings.MaxConcurrentDocuments);
            Assert.Equal(_appSettings.Api.BaseUrl, _viewModel.ValidationSettings.ApiBaseUrl);
            Assert.Equal((int)_appSettings.Api.Timeout.TotalSeconds, _viewModel.ValidationSettings.ApiTimeoutSeconds);
            Assert.Equal(_appSettings.Api.EnableCaching, _viewModel.ValidationSettings.EnableApiCaching);
            Assert.Equal((int)_appSettings.Api.CacheExpiry.TotalHours, _viewModel.ValidationSettings.ApiCacheExpiryHours);
        }

        [Fact]
        public void UpdateApiSettings_ShouldUpdatePropertiesCorrectly()
        {
            // Arrange
            var newApiUrl = "https://api.example.com/test";
            var newApiKey = "test-api-key-123";
            var newTimeout = 60;
            var newCacheExpiry = 4;

            // Act
            _viewModel.ValidationSettings.ApiBaseUrl = newApiUrl;
            _viewModel.ValidationSettings.ApiKey = newApiKey;
            _viewModel.ValidationSettings.ApiTimeoutSeconds = newTimeout;
            _viewModel.ValidationSettings.EnableApiCaching = false;
            _viewModel.ValidationSettings.ApiCacheExpiryHours = newCacheExpiry;

            // Assert
            Assert.Equal(newApiUrl, _viewModel.ValidationSettings.ApiBaseUrl);
            Assert.Equal(newApiKey, _viewModel.ValidationSettings.ApiKey);
            Assert.Equal(newTimeout, _viewModel.ValidationSettings.ApiTimeoutSeconds);
            Assert.False(_viewModel.ValidationSettings.EnableApiCaching);
            Assert.Equal(newCacheExpiry, _viewModel.ValidationSettings.ApiCacheExpiryHours);
        }

        [Fact]
        public void ValidateSettings_WithValidApiUrl_ShouldReturnTrue()
        {
            // Arrange
            _viewModel.ValidationSettings.ApiBaseUrl = "https://api.example.com";
            _viewModel.ValidationSettings.ApiTimeoutSeconds = 30;
            _viewModel.ValidationSettings.ApiCacheExpiryHours = 2;

            // Act
            var result = InvokeValidateSettings();

            // Assert
            Assert.True(result);
        }

        [Fact]
        public void ValidateSettings_WithTestApiUrl_ShouldReturnTrue()
        {
            // Arrange
            _viewModel.ValidationSettings.ApiBaseUrl = "Test";
            _viewModel.ValidationSettings.ApiTimeoutSeconds = 30;
            _viewModel.ValidationSettings.ApiCacheExpiryHours = 2;

            // Act
            var result = InvokeValidateSettings();

            // Assert
            Assert.True(result);
        }

        [Theory]
        [InlineData("invalid-url")]
        [InlineData("ftp://invalid.com")]
        [InlineData("not-a-url")]
        public void ValidateSettings_WithInvalidApiUrl_ShouldReturnFalse(string invalidUrl)
        {
            // Arrange
            _viewModel.ValidationSettings.ApiBaseUrl = invalidUrl;
            _viewModel.ValidationSettings.ApiTimeoutSeconds = 30;
            _viewModel.ValidationSettings.ApiCacheExpiryHours = 2;

            // Act
            var result = InvokeValidateSettings();

            // Assert
            Assert.False(result);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(301)]
        [InlineData(-1)]
        public void ValidateSettings_WithInvalidApiTimeout_ShouldReturnFalse(int invalidTimeout)
        {
            // Arrange
            _viewModel.ValidationSettings.ApiBaseUrl = "https://api.example.com";
            _viewModel.ValidationSettings.ApiTimeoutSeconds = invalidTimeout;
            _viewModel.ValidationSettings.ApiCacheExpiryHours = 2;

            // Act
            var result = InvokeValidateSettings();

            // Assert
            Assert.False(result);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(25)]
        [InlineData(-1)]
        public void ValidateSettings_WithInvalidCacheExpiry_ShouldReturnFalse(int invalidCacheExpiry)
        {
            // Arrange
            _viewModel.ValidationSettings.ApiBaseUrl = "https://api.example.com";
            _viewModel.ValidationSettings.ApiTimeoutSeconds = 30;
            _viewModel.ValidationSettings.ApiCacheExpiryHours = invalidCacheExpiry;

            // Act
            var result = InvokeValidateSettings();

            // Assert
            Assert.False(result);
        }

        [Fact]
        public async Task TestApiConnection_WithTestUrl_ShouldSucceed()
        {
            // Arrange
            _viewModel.ValidationSettings.ApiBaseUrl = "Test";

            // Act
            await _viewModel.ValidationSettings.TestApiConnectionCommand.ExecuteAsync(null);

            // Assert
            _mockLogger.Verify(
                x => x.LogInformation("API connection test successful (Test mode)"),
                Times.Once);
        }

        [Fact]
        public async Task TestApiConnection_WithEmptyUrl_ShouldLogWarning()
        {
            // Arrange
            _viewModel.ValidationSettings.ApiBaseUrl = "";

            // Act
            await _viewModel.ValidationSettings.TestApiConnectionCommand.ExecuteAsync(null);

            // Assert
            _mockLogger.Verify(
                x => x.LogWarning("API Base URL is required for connection test"),
                Times.Once);
        }

        private static AppSettings CreateTestAppSettings()
        {
            return new AppSettings
            {
                Processing = new ProcessingSettings
                {
                    MaxConcurrentDocuments = 10,
                    BatchSize = 50,
                    LookupIdPattern = @"(TSRC-[^-]+-[0-9]{6}|CMS-[^-]+-[0-9]{6})"
                },
                Validation = new ValidationSettings
                {
                    HttpTimeout = System.TimeSpan.FromSeconds(30),
                    MaxRetryAttempts = 3,
                    UserAgent = "BulkEditor/1.0"
                },
                Backup = new BackupSettings
                {
                    BackupDirectory = "Backups",
                    MaxBackupAge = 30
                },
                Logging = new LoggingSettings
                {
                    LogLevel = "Information",
                    LogDirectory = "Logs"
                },
                Replacement = new ReplacementSettings
                {
                    MaxReplacementRules = 50
                },
                Api = new ApiSettings
                {
                    BaseUrl = "https://api.example.com",
                    ApiKey = "test-key",
                    Timeout = System.TimeSpan.FromSeconds(30),
                    EnableCaching = true,
                    CacheExpiry = System.TimeSpan.FromHours(2)
                },
                Update = new UpdateSettings
                {
                    AutoUpdateEnabled = true,
                    CheckIntervalHours = 24,
                    InstallSecurityUpdatesAutomatically = true,
                    NotifyOnUpdatesAvailable = true,
                    CreateBackupBeforeUpdate = true,
                    IncludePrerelease = false
                }
            };
        }

        private bool InvokeValidateSettings()
        {
            // Use reflection to access the private ValidateSettings method
            var method = typeof(SettingsViewModel).GetMethod("ValidateSettings",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            return (bool)method!.Invoke(_viewModel, null)!;
        }
    }
}