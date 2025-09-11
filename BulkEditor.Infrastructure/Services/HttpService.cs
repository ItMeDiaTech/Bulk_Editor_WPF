using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using System;
using System.Net;
using System.Net.Http;
using System.Text.Json;
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
        private readonly IRetryPolicyService _retryPolicyService;
        private readonly IStructuredLoggingService _structuredLogger;

        public HttpService(HttpClient httpClient, ILoggingService logger, IRetryPolicyService retryPolicyService, IStructuredLoggingService structuredLogger)
        {
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _retryPolicyService = retryPolicyService ?? throw new ArgumentNullException(nameof(retryPolicyService));
            _structuredLogger = structuredLogger ?? throw new ArgumentNullException(nameof(structuredLogger));

            // Note: HttpClient configuration moved to DI container to avoid modification after first request
        }

        public async Task<HttpResponseMessage> GetAsync(string url, CancellationToken cancellationToken = default)
        {
            var correlationId = Guid.NewGuid().ToString();
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            var retryPolicy = _retryPolicyService.CreateHttpRetryPolicy();
            
            return await _retryPolicyService.ExecuteWithRetryAsync(async () =>
            {
                // Log structured HTTP request
                var requestEntry = new HttpRequestLogEntry
                {
                    Method = "GET",
                    Url = url,
                    CorrelationId = correlationId,
                    OperationName = "HttpGet",
                    UserAgent = _httpClient.DefaultRequestHeaders.UserAgent?.ToString()
                };

                await _structuredLogger.LogHttpRequestAsync(requestEntry);

                var response = await _httpClient.GetAsync(url, cancellationToken).ConfigureAwait(false);
                
                stopwatch.Stop();

                // Log structured HTTP response
                var responseEntry = new HttpResponseLogEntry
                {
                    StatusCode = (int)response.StatusCode,
                    StatusDescription = response.ReasonPhrase ?? "Unknown",
                    Duration = stopwatch.Elapsed,
                    CorrelationId = correlationId,
                    OperationName = "HttpGet",
                    IsSuccessStatusCode = response.IsSuccessStatusCode,
                    ContentLength = response.Content.Headers.ContentLength ?? 0
                };

                await _structuredLogger.LogHttpResponseAsync(responseEntry);

                return response;
            }, retryPolicy, cancellationToken).ConfigureAwait(false);
        }

        public async Task<HttpResponseMessage> PostJsonAsync(string url, object data, CancellationToken cancellationToken = default)
        {
            // Handle test mode immediately without retry logic
            if (url.ToLower() == "test")
            {
                return CreateTestApiResponse();
            }

            var correlationId = Guid.NewGuid().ToString();
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            var retryPolicy = _retryPolicyService.CreateHttpRetryPolicy();
            
            return await _retryPolicyService.ExecuteWithRetryAsync(async () =>
            {
                // CRITICAL FIX: Ensure JSON matches VBA API expectations (Issue #4)
                var json = System.Text.Json.JsonSerializer.Serialize(data, new JsonSerializerOptions
                {
                    PropertyNamingPolicy = null, // Keep exact property names like "Lookup_ID"
                    WriteIndented = false
                });

                // Log structured HTTP request
                var requestEntry = new HttpRequestLogEntry
                {
                    Method = "POST",
                    Url = url,
                    Body = json,
                    ContentLength = System.Text.Encoding.UTF8.GetByteCount(json),
                    CorrelationId = correlationId,
                    OperationName = "ApiLookupRequest",
                    UserAgent = _httpClient.DefaultRequestHeaders.UserAgent?.ToString(),
                    Headers = new Dictionary<string, string>
                    {
                        ["Content-Type"] = "application/json"
                    }
                };

                await _structuredLogger.LogHttpRequestAsync(requestEntry);

                var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync(url, content, cancellationToken).ConfigureAwait(false);
                
                stopwatch.Stop();

                // Read response content for logging
                var responseContent = string.Empty;
                if (response.Content != null)
                {
                    responseContent = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                }

                // Log structured HTTP response
                var responseEntry = new HttpResponseLogEntry
                {
                    StatusCode = (int)response.StatusCode,
                    StatusDescription = response.ReasonPhrase ?? "Unknown",
                    Body = responseContent,
                    ContentLength = responseContent?.Length ?? 0,
                    Duration = stopwatch.Elapsed,
                    CorrelationId = correlationId,
                    OperationName = "ApiLookupRequest",
                    IsSuccessStatusCode = response.IsSuccessStatusCode
                };

                await _structuredLogger.LogHttpResponseAsync(responseEntry);

                return response;
            }, retryPolicy, cancellationToken).ConfigureAwait(false);
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
                        Title = "🔄 UPDATED API Title for Testing - V2.0",
                        Status = "Released"
                    },
                    new
                    {
                        Document_ID = "test-doc-002",
                        Content_ID = "TEST-CONTENT-789012",
                        Title = "🔄 UPDATED Expired Document Title - V2.0",
                        Status = "Expired"
                    },
                    new
                    {
                        Document_ID = "test-doc-003",
                        Content_ID = "TEST-CONTENT-345678",
                        Title = "🔄 UPDATED Released Document Title - V2.0",
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
            // Handle file:// URLs which are not supported by HttpClient
            if (url.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
            {
                _logger.LogDebug("Handling file:// URL: {Url}", url);

                // Extract local file path from file:// URL
                var filePath = new Uri(url).LocalPath;
                var fileExists = System.IO.File.Exists(filePath);

                var fileResponse = new HttpResponseMessage(fileExists ? HttpStatusCode.OK : HttpStatusCode.NotFound);
                _logger.LogDebug("File {FilePath} exists: {FileExists}", filePath, fileExists);
                return fileResponse;
            }

            var retryPolicy = _retryPolicyService.CreateHttpRetryPolicy();
            
            return await _retryPolicyService.ExecuteWithRetryAsync(async () =>
            {
                _logger.LogDebug("Sending HEAD request to: {Url}", url);
                var request = new HttpRequestMessage(HttpMethod.Head, url);
                var response = await _httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
                _logger.LogDebug("Received response {StatusCode} from HEAD request: {Url}", response.StatusCode, url);
                return response;
            }, retryPolicy, cancellationToken).ConfigureAwait(false);
        }

        public async Task<bool> IsUrlAccessibleAsync(string url, CancellationToken cancellationToken = default)
        {
            try
            {
                // Handle file:// URLs directly
                if (url.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                {
                    var filePath = new Uri(url).LocalPath;
                    var fileAccessible = System.IO.File.Exists(filePath);
                    _logger.LogDebug("File URL {Url} accessibility check: {IsAccessible}", url, fileAccessible);
                    return fileAccessible;
                }

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
                // Handle file:// URLs directly
                if (url.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                {
                    var filePath = new Uri(url).LocalPath;
                    var fileExists = System.IO.File.Exists(filePath);
                    var actualStatus = fileExists ? HttpStatusCode.OK : HttpStatusCode.NotFound;
                    var statusMatches = actualStatus == expectedStatus;
                    _logger.LogDebug("File URL {Url} status check: Expected {ExpectedStatus}, Got {ActualStatus}, Match: {HasExpectedStatus}",
                        url, expectedStatus, actualStatus, statusMatches);
                    return statusMatches;
                }

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

        public async Task<bool> TestConnectionAsync(string url, string apiKey = "")
        {
            try
            {
                var request = new HttpRequestMessage(HttpMethod.Head, url);
                if (!string.IsNullOrWhiteSpace(apiKey))
                {
                    request.Headers.Add("Authorization", $"Bearer {apiKey}");
                }

                var response = await _httpClient.SendAsync(request);
                return response.IsSuccessStatusCode;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to test connection to {Url}", url);
                return false;
            }
        }
    }
}

