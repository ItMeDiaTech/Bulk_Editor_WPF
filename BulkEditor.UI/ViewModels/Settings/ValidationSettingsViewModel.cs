using BulkEditor.Core.Interfaces;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Net.Http;
using System.Threading.Tasks;

namespace BulkEditor.UI.ViewModels.Settings
{
    public partial class ValidationSettingsViewModel : ObservableObject
    {
        private readonly ILoggingService? _logger;
        private readonly IHttpService? _httpService;

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

        [ObservableProperty]
        private bool _isBusy;

        [ObservableProperty]
        private string _busyMessage = string.Empty;

        public ValidationSettingsViewModel(ILoggingService? logger = null, IHttpService? httpService = null)
        {
            _logger = logger;
            _httpService = httpService;
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
                    _logger?.LogWarning("API Base URL is required for connection test");
                    System.Windows.MessageBox.Show(
                        "API Base URL is required for connection test. Please enter a valid URL.",
                        "API Test Failed",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                    return;
                }

                if (ApiBaseUrl.ToLower() == "test")
                {
                    _logger?.LogInformation("API connection test successful (Test mode)");
                    System.Windows.MessageBox.Show(
                        "API connection test successful!\n\nTest mode is active - using mock API responses.",
                        "API Test Successful",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                    return;
                }

                if (_httpService == null)
                {
                    _logger?.LogWarning("HTTP service is not available for API connection test");
                    System.Windows.MessageBox.Show(
                        "HTTP service is not available. Cannot test API connection.",
                        "API Test Failed",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Error);
                    return;
                }

                var success = await _httpService.TestConnectionAsync(ApiBaseUrl, ApiKey);

                if (success)
                {
                    _logger?.LogInformation("API connection test successful");
                    System.Windows.MessageBox.Show(
                        "API connection test successful!",
                        "API Test Successful",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                }
                else
                {
                    _logger?.LogWarning("API connection test failed");
                    System.Windows.MessageBox.Show(
                        "API connection test failed. Please check the URL and your internet connection.",
                        "API Test Failed",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Error);
                }
            }
            catch (TaskCanceledException ex) when (ex.InnerException is TimeoutException)
            {
                _logger?.LogWarning("API connection test timed out after {TimeoutSeconds} seconds", ApiTimeoutSeconds);

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
                _logger?.LogError(ex, "API connection test failed with HTTP error");

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
                _logger?.LogError(ex, "API connection test failed with unexpected error");

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
                BusyMessage = string.Empty;
            }
        }
    }
}