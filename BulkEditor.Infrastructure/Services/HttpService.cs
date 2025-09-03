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

            // Note: HttpClient configuration moved to DI container to avoid modification after first request
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

        public async Task<HttpResponseMessage> PostJsonAsync(string url, object data, CancellationToken cancellationToken = default)
        {
            try
            {
                _logger.LogDebug("Sending POST request to: {Url}", url);

                // Handle test mode
                if (url.ToLower() == "test")
                {
                    return CreateTestApiResponse();
                }

                var json = System.Text.Json.JsonSerializer.Serialize(data);
                var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync(url, content, cancellationToken);
                _logger.LogDebug("Received response {StatusCode} from POST: {Url}", response.StatusCode, url);
                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error sending POST request to: {Url}", url);
                throw;
            }
        }

        private HttpResponseMessage CreateTestApiResponse()
        {
            var testResponse = new
            {
                Version = "2.1",
                Changes = "Test mode - mock data response",
                Results = new[]
                {
                    new
                    {
                        Document_ID = "test-doc-001",
                        Content_ID = "TEST-CONTENT-123456",
                        Title = "Test Document Title",
                        Status = "Released"
                    },
                    new
                    {
                        Document_ID = "test-doc-002",
                        Content_ID = "TEST-CONTENT-789012",
                        Title = "Expired Test Document",
                        Status = "Expired"
                    },
                    new
                    {
                        Document_ID = "test-doc-003",
                        Content_ID = "TEST-CONTENT-345678",
                        Title = "Another Released Document",
                        Status = "Released"
                    }
                }
            };

            var json = System.Text.Json.JsonSerializer.Serialize(testResponse);
            var response = new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json")
            };

            _logger.LogInformation("Generated test API response with {ResultCount} results", testResponse.Results.Length);
            return response;
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
            // Note: Timeout modification disabled to prevent "Properties can only be modified before sending the first request" error
            _logger.LogDebug("HTTP client timeout change requested: {Timeout} (not applied to avoid HttpClient configuration errors)", timeout);
        }

        public void SetUserAgent(string userAgent)
        {
            // Note: User-Agent modification disabled to prevent "Properties can only be modified before sending the first request" error
            _logger.LogDebug("HTTP client user agent change requested: {UserAgent} (not applied to avoid HttpClient configuration errors)", userAgent);
        }
    }
}