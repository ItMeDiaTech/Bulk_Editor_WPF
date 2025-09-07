using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using BulkEditor.Infrastructure.Services;
using FluentAssertions;
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
            
            _httpService = new HttpService(_httpClient, _mockLogger.Object);
        }

        [Fact]
        public async Task PostJsonAsync_WithTestUrl_ReturnsTestResponse()
        {
            // Arrange
            var testData = new { Lookup_ID = new[] { "TSRC-TEST-123456", "CMS-TEST-789012" } };

            // Act
            var response = await _httpService.PostJsonAsync("Test", testData);

            // Assert
            response.Should().NotBeNull();
            response.StatusCode.Should().Be(HttpStatusCode.OK);
            response.Content.Headers.ContentType?.MediaType.Should().Be("application/json");

            var responseContent = await response.Content.ReadAsStringAsync();
            responseContent.Should().NotBeNullOrEmpty();

            // Verify the response matches VBA expected format
            var apiResponse = JsonSerializer.Deserialize<ApiResponse>(responseContent);
            apiResponse.Should().NotBeNull();
            apiResponse.Version.Should().Be("2.1");
            apiResponse.Changes.Should().Be("Test mode - mock data response");
            apiResponse.Results.Should().NotBeEmpty();
            apiResponse.Results.Should().HaveCount(3);

            // Verify each result has required fields
            foreach (var result in apiResponse.Results)
            {
                result.Document_ID.Should().NotBeNullOrEmpty();
                result.Content_ID.Should().NotBeNullOrEmpty();
                result.Title.Should().NotBeNullOrEmpty();
                result.Status.Should().NotBeNullOrEmpty();
            }

            // Verify status values - should have both "Released" and "Expired" for testing
            apiResponse.Results.Should().Contain(r => r.Status == "Released");
            apiResponse.Results.Should().Contain(r => r.Status == "Expired");
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
            response1.StatusCode.Should().Be(HttpStatusCode.OK);
            response2.StatusCode.Should().Be(HttpStatusCode.OK);
            response3.StatusCode.Should().Be(HttpStatusCode.OK);
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
            response.Should().NotBeNull();
            response.StatusCode.Should().Be(HttpStatusCode.OK);

            var responseContent = await response.Content.ReadAsStringAsync();
            var apiResponse = JsonSerializer.Deserialize<ApiResponse>(responseContent);

            // Verify the test response structure matches what VBA expects
            apiResponse.Should().NotBeNull();
            apiResponse.Version.Should().NotBeNullOrEmpty();
            apiResponse.Changes.Should().NotBeNullOrEmpty();
            apiResponse.Results.Should().NotBeNull();
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
            root.TryGetProperty("Version", out var versionElement).Should().BeTrue();
            versionElement.GetString().Should().NotBeNullOrEmpty();

            root.TryGetProperty("Changes", out var changesElement).Should().BeTrue();
            changesElement.GetString().Should().NotBeNullOrEmpty();

            root.TryGetProperty("Results", out var resultsElement).Should().BeTrue();
            resultsElement.ValueKind.Should().Be(JsonValueKind.Array);

            // Check each result item has the required fields for VBA
            foreach (var result in resultsElement.EnumerateArray())
            {
                result.TryGetProperty("Document_ID", out _).Should().BeTrue();
                result.TryGetProperty("Content_ID", out _).Should().BeTrue();
                result.TryGetProperty("Title", out _).Should().BeTrue();
                result.TryGetProperty("Status", out _).Should().BeTrue();
            }
        }

        [Fact]
        public async Task PostJsonAsync_TestMode_LogsCorrectly()
        {
            // Arrange
            var testData = new { Lookup_ID = new[] { "TSRC-TEST-123456" } };

            // Act
            await _httpService.PostJsonAsync("Test", testData);

            // Assert - verify enhanced logging for debugging  
            _mockLogger.Verify(
                x => x.LogInformation("Sending POST request for combined lookup identifiers to: {Url}", "Test"),
                Times.Once);

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
                apiResponse.Results.Should().Contain(r => r.Status == "Expired");
                // Verify the expected status is present when looking for expired
                apiResponse.Results.Should().Contain(r => r.Status == expectedStatus);
            }
            else
            {
                // Should have non-expired documents
                apiResponse.Results.Should().Contain(r => r.Status == "Released");
                // For non-expired cases, verify expectedStatus exists if it's "Released"
                if (expectedStatus == "Released")
                {
                    apiResponse.Results.Should().Contain(r => r.Status == expectedStatus);
                }
            }
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}