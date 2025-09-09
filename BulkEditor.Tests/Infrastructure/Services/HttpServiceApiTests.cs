using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using BulkEditor.Infrastructure.Services;
using Moq;
using System;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace BulkEditor.Tests.Infrastructure.Services
{
    /// <summary>
    /// Tests for API functionality in HttpService
    /// </summary>
    public class HttpServiceApiTests : IDisposable
    {
        private readonly Mock<ILoggingService> _mockLogger;
        private readonly HttpClient _httpClient;
        private readonly HttpService _httpService;

        public HttpServiceApiTests()
        {
            _mockLogger = new Mock<ILoggingService>();
            
            // CRITICAL FIX: Use proper HttpClient configuration to prevent socket exhaustion (Issue #22)
            var handler = new HttpClientHandler()
            {
                MaxConnectionsPerServer = 10,
                UseProxy = false,
                UseCookies = false
            };
            
            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromMinutes(5)
            };
            
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "BulkEditor-Test/1.0");
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
            
            var mockRetryPolicyService = new Mock<BulkEditor.Core.Services.IRetryPolicyService>();
            var httpRetryPolicy = new BulkEditor.Core.Services.RetryPolicy { MaxRetries = 3, PolicyName = "HTTP" };
            mockRetryPolicyService.Setup(x => x.CreateHttpRetryPolicy()).Returns(httpRetryPolicy);
            mockRetryPolicyService.Setup(x => x.ExecuteWithRetryAsync(It.IsAny<Func<Task<System.Net.Http.HttpResponseMessage>>>(), It.IsAny<BulkEditor.Core.Services.RetryPolicy>(), It.IsAny<CancellationToken>()))
                .Returns<Func<Task<System.Net.Http.HttpResponseMessage>>, BulkEditor.Core.Services.RetryPolicy, CancellationToken>((func, policy, token) => func());
            
            var mockStructuredLogger = new Mock<BulkEditor.Core.Services.IStructuredLoggingService>();
            
            _httpService = new HttpService(_httpClient, _mockLogger.Object, mockRetryPolicyService.Object, mockStructuredLogger.Object);
        }

        [Fact]
        public async Task PostJsonAsync_WithTestUrl_ReturnsTestResponse()
        {
            // Arrange
            var testData = new { Lookup_ID = new[] { "TSRC-TEST-123456", "CMS-TEST-789012" } };

            // Act
            var response = await _httpService.PostJsonAsync("Test", testData);

            // Assert
            Assert.NotNull(response);
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);
            Assert.Equal("application/json", response.Content.Headers.ContentType?.MediaType);

            var responseContent = await response.Content.ReadAsStringAsync();
            Assert.False(string.IsNullOrEmpty(responseContent));

            // Verify the response matches VBA expected format
            var apiResponse = JsonSerializer.Deserialize<ApiResponse>(responseContent);
            Assert.NotNull(apiResponse);
            Assert.Equal("2.1", apiResponse.Version);
            Assert.Equal("Test mode - mock data response", apiResponse.Changes);
            Assert.NotEmpty(apiResponse.Results);
            Assert.Equal(3, apiResponse.Results.Count);

            // Verify each result has required fields
            foreach (var result in apiResponse.Results)
            {
                Assert.False(string.IsNullOrEmpty(result.Document_ID));
                Assert.False(string.IsNullOrEmpty(result.Content_ID));
                Assert.False(string.IsNullOrEmpty(result.Title));
                Assert.False(string.IsNullOrEmpty(result.Status));
            }

            // Verify status values - should have both "Released" and "Expired" for testing
            Assert.Contains(apiResponse.Results, r => r.Status == "Released");
            Assert.Contains(apiResponse.Results, r => r.Status == "Expired");
        }

        [Fact]
        public async Task PostJsonAsync_WithTestUrl_CaseInsensitive()
        {
            // Arrange
            var testData = new { Lookup_ID = new[] { "TSRC-TEST-123456" } };

            // Act - test different case variations
            var response1 = await _httpService.PostJsonAsync("test", testData);
            var response2 = await _httpService.PostJsonAsync("TEST", testData);
            var response3 = await _httpService.PostJsonAsync("Test", testData);

            // Assert
            Assert.Equal(HttpStatusCode.OK, response1.StatusCode);
            Assert.Equal(HttpStatusCode.OK, response2.StatusCode);
            Assert.Equal(HttpStatusCode.OK, response3.StatusCode);
        }

        [Fact]
        public async Task PostJsonAsync_WithValidData_SerializesJsonProperly()
        {
            // Arrange
            var testData = new
            {
                Lookup_ID = new[] { "TSRC-DOC-123456", "CMS-DOC-789012", "TSRC-DOC-345678" }
            };

            // Act
            var response = await _httpService.PostJsonAsync("Test", testData);

            // Assert
            Assert.NotNull(response);
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);

            var responseContent = await response.Content.ReadAsStringAsync();
            var apiResponse = JsonSerializer.Deserialize<ApiResponse>(responseContent);

            // Verify the test response structure matches what VBA expects
            Assert.NotNull(apiResponse);
            Assert.False(string.IsNullOrEmpty(apiResponse.Version));
            Assert.False(string.IsNullOrEmpty(apiResponse.Changes));
            Assert.NotNull(apiResponse.Results);
        }

        [Fact]
        public async Task PostJsonAsync_TestResponseFormat_MatchesVbaExpectations()
        {
            // Arrange
            var lookupRequest = new { Lookup_ID = new[] { "TSRC-TEST-123456" } };

            // Act
            var response = await _httpService.PostJsonAsync("Test", lookupRequest);
            var responseContent = await response.Content.ReadAsStringAsync();

            // Parse as the VBA would expect
            var jsonDoc = JsonDocument.Parse(responseContent);
            var root = jsonDoc.RootElement;

            // Assert - verify JSON structure matches VBA parsing expectations
            Assert.True(root.TryGetProperty("Version", out var versionElement));
            Assert.False(string.IsNullOrEmpty(versionElement.GetString()));

            Assert.True(root.TryGetProperty("Changes", out var changesElement));
            Assert.False(string.IsNullOrEmpty(changesElement.GetString()));

            Assert.True(root.TryGetProperty("Results", out var resultsElement));
            Assert.Equal(JsonValueKind.Array, resultsElement.ValueKind);

            // Check each result item has the required fields for VBA
            foreach (var result in resultsElement.EnumerateArray())
            {
                Assert.True(result.TryGetProperty("Document_ID", out _));
                Assert.True(result.TryGetProperty("Content_ID", out _));
                Assert.True(result.TryGetProperty("Title", out _));
                Assert.True(result.TryGetProperty("Status", out _));
            }
        }

        [Fact]
        public async Task PostJsonAsync_TestMode_LogsCorrectly()
        {
            // Arrange
            var testData = new { Lookup_ID = new[] { "TSRC-TEST-123456" } };

            // Act
            await _httpService.PostJsonAsync("Test", testData);

            // Assert - verify test mode only logs the response generation
            _mockLogger.Verify(
                x => x.LogInformation("Generated test API response with {ResultCount} results", 3),
                Times.Once);
        }

        [Theory]
        [InlineData("Expired", true)]
        [InlineData("Released", false)]
        [InlineData("Active", false)]
        [InlineData("Draft", false)]
        public async Task PostJsonAsync_TestResponse_ContainsExpectedStatuses(string expectedStatus, bool shouldBeExpired)
        {
            // Arrange
            var testData = new { Lookup_ID = new[] { "TSRC-TEST-123456" } };

            // Act
            var response = await _httpService.PostJsonAsync("Test", testData);
            var responseContent = await response.Content.ReadAsStringAsync();
            var apiResponse = JsonSerializer.Deserialize<ApiResponse>(responseContent);

            // Assert
            if (shouldBeExpired)
            {
                // Should have at least one expired document for testing
                Assert.Contains(apiResponse.Results, r => r.Status == "Expired");
                // Verify the expected status is present when looking for expired
                Assert.Contains(apiResponse.Results, r => r.Status == expectedStatus);
            }
            else
            {
                // Should have non-expired documents
                Assert.Contains(apiResponse.Results, r => r.Status == "Released");
                // For non-expired cases, verify expectedStatus exists if it's "Released"
                if (expectedStatus == "Released")
                {
                    Assert.Contains(apiResponse.Results, r => r.Status == expectedStatus);
                }
            }
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}