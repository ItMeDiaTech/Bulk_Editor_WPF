using BulkEditor.Core.Interfaces;
using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Implementation of HTTP service for web requests
    /// </summary>
    public class HttpService : IHttpService
    {
        private readonly HttpClient _httpClient;
        private readonly ILoggingService _logger;

        public HttpService(HttpClient httpClient, ILoggingService logger)
        {
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));

            // Set default timeout and user agent
            _httpClient.Timeout = TimeSpan.FromSeconds(30);
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "BulkEditor/1.0");
        }

        public async Task<HttpResponseMessage> GetAsync(string url, CancellationToken cancellationToken = default)
        {
            try
            {
                _logger.LogDebug("Sending GET request to: {Url}", url);
                var response = await _httpClient.GetAsync(url, cancellationToken);
                _logger.LogDebug("Received response {StatusCode} from: {Url}", response.StatusCode, url);
                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error sending GET request to: {Url}", url);
                throw;
            }
        }

        public async Task<HttpResponseMessage> HeadAsync(string url, CancellationToken cancellationToken = default)
        {
            try
            {
                _logger.LogDebug("Sending HEAD request to: {Url}", url);
                var request = new HttpRequestMessage(HttpMethod.Head, url);
                var response = await _httpClient.SendAsync(request, cancellationToken);
                _logger.LogDebug("Received response {StatusCode} from HEAD request: {Url}", response.StatusCode, url);
                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error sending HEAD request to: {Url}", url);
                throw;
            }
        }

        public async Task<bool> IsUrlAccessibleAsync(string url, CancellationToken cancellationToken = default)
        {
            try
            {
                using var response = await HeadAsync(url, cancellationToken);
                var isAccessible = response.IsSuccessStatusCode;
                _logger.LogDebug("URL {Url} accessibility check: {IsAccessible} (Status: {StatusCode})",
                    url, isAccessible, response.StatusCode);
                return isAccessible;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("URL accessibility check failed for: {Url}. Error: {Error}", url, ex.Message);
                return false;
            }
        }

        public async Task<string> GetContentAsync(string url, CancellationToken cancellationToken = default)
        {
            try
            {
                _logger.LogDebug("Getting content from: {Url}", url);
                using var response = await GetAsync(url, cancellationToken);
                response.EnsureSuccessStatusCode();

                var content = await response.Content.ReadAsStringAsync(cancellationToken);
                _logger.LogDebug("Retrieved {ContentLength} characters from: {Url}", content.Length, url);
                return content;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting content from: {Url}", url);
                throw;
            }
        }

        public async Task<bool> CheckUrlStatusAsync(string url, HttpStatusCode expectedStatus, CancellationToken cancellationToken = default)
        {
            try
            {
                using var response = await HeadAsync(url, cancellationToken);
                var hasExpectedStatus = response.StatusCode == expectedStatus;
                _logger.LogDebug("URL {Url} status check: Expected {ExpectedStatus}, Got {ActualStatus}, Match: {HasExpectedStatus}",
                    url, expectedStatus, response.StatusCode, hasExpectedStatus);
                return hasExpectedStatus;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("URL status check failed for: {Url}. Error: {Error}", url, ex.Message);
                return false;
            }
        }

        public void SetTimeout(TimeSpan timeout)
        {
            _httpClient.Timeout = timeout;
            _logger.LogDebug("HTTP client timeout set to: {Timeout}", timeout);
        }

        public void SetUserAgent(string userAgent)
        {
            _httpClient.DefaultRequestHeaders.Remove("User-Agent");
            _httpClient.DefaultRequestHeaders.Add("User-Agent", userAgent);
            _logger.LogDebug("HTTP client user agent set to: {UserAgent}", userAgent);
        }
    }
}