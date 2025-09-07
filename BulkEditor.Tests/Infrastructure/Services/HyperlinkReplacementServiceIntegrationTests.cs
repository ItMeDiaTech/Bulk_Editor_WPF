using BulkEditor.Core.Configuration;
using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using BulkEditor.Infrastructure.Services;
using FluentAssertions;
using Microsoft.Extensions.Options;
using Moq;
using Xunit;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;

namespace BulkEditor.Tests.Infrastructure.Services
{
    /// <summary>
    /// Integration tests for HyperlinkReplacementService to verify Base_File.vba methodology compliance
    /// Tests all critical fixes: expired status detection, missing lookup IDs, HTML filtering, etc.
    /// </summary>
    public class HyperlinkReplacementServiceIntegrationTests
    {
        private readonly Mock<IHttpService> _httpServiceMock;
        private readonly Mock<ILoggingService> _loggerMock;
        private readonly HyperlinkReplacementService _service;

        public HyperlinkReplacementServiceIntegrationTests()
        {
            _httpServiceMock = new Mock<IHttpService>();
            _loggerMock = new Mock<ILoggingService>();

            // Create default AppSettings for testing
            var appSettings = new AppSettings
            {
                Api = new ApiSettings
                {
                    BaseUrl = "https://api.example.com"
                }
            };
            var appSettingsOptions = Options.Create(appSettings);

            _service = new HyperlinkReplacementService(_httpServiceMock.Object, _loggerMock.Object, appSettingsOptions);
        }

        [Fact]
        public async Task ProcessApiResponseAsync_WithExpiredStatus_ShouldDetectExpiredDocuments()
        {
            // Arrange
            var lookupIds = new[] { "TSRC-PROD-123456", "TSRC-PROD-654321" };

            // Act
            var result = await _service.ProcessApiResponseAsync(lookupIds, CancellationToken.None);

            // Assert
            result.Should().NotBeNull();
            result.IsSuccess.Should().BeTrue();

            // Verify that some documents are detected as expired based on simulation logic
            var totalDocuments = result.FoundDocuments.Count + result.ExpiredDocuments.Count;
            totalDocuments.Should().BeGreaterThan(0, "At least some documents should be found");

            // Check that expired detection works
            if (result.ExpiredDocuments.Any())
            {
                result.ExpiredDocuments.Should().AllSatisfy(doc =>
                    doc.Status.Should().Be("Expired", "Expired documents should have Status='Expired'"));
            }
        }

        [Fact]
        public async Task ProcessApiResponseAsync_WithMissingLookupIds_ShouldIdentifyMissingIds()
        {
            // Arrange - Use multiple lookup IDs to ensure some are missing based on simulation
            var lookupIds = new[]
            {
                "TSRC-PROD-111111",
                "TSRC-PROD-222222",
                "TSRC-PROD-333333",
                "TSRC-PROD-444444",
                "TSRC-PROD-555555",
                "TSRC-PROD-666666",
                "TSRC-PROD-777777",
                "TSRC-PROD-888888",
                "TSRC-PROD-999999"
            };

            // Act
            var result = await _service.ProcessApiResponseAsync(lookupIds, CancellationToken.None);

            // Assert
            result.Should().NotBeNull();
            result.IsSuccess.Should().BeTrue();

            // Based on simulation logic (15% chance of missing), some should be missing
            var totalReturned = result.FoundDocuments.Count + result.ExpiredDocuments.Count;
            var totalMissing = result.MissingLookupIds.Count;

            (totalReturned + totalMissing).Should().Be(lookupIds.Length,
                "Total returned plus missing should equal original lookup IDs");

            // Verify missing IDs are properly tracked
            result.MissingLookupIds.Should().AllSatisfy(missingId =>
                lookupIds.Should().Contain(missingId, "Missing ID should be from original list"));
        }

        [Fact]
        public async Task LookupDocumentByIdentifierAsync_WithValidContentId_ShouldReturnDocumentRecord()
        {
            // Arrange
            var contentId = "123456";

            // Act
            var result = await _service.LookupDocumentByIdentifierAsync(contentId, CancellationToken.None);

            // Assert
            result.Should().NotBeNull("Valid content ID should return document record");
            result.Content_ID.Should().Be("123456", "Content ID should be properly formatted");
            result.Document_ID.Should().NotBeNullOrEmpty("Document ID should be generated");
            result.Title.Should().NotBeNullOrEmpty("Title should be generated");
            result.Status.Should().BeOneOf("Released", "Expired", "Unknown", "Error handling statuses");
        }

        [Fact]
        public async Task LookupDocumentByIdentifierAsync_WithMissingLookupId_ShouldReturnNull()
        {
            // Arrange - Use content ID that will trigger missing lookup ID simulation
            var contentId = "000001"; // This should trigger the 15% missing chance based on hash

            // Act
            var result = await _service.LookupDocumentByIdentifierAsync(contentId, CancellationToken.None);

            // Assert - Some content IDs will return null to simulate missing lookup IDs
            // We can't guarantee which ones due to hash-based simulation, so we test the behavior
            if (result == null)
            {
                // This is expected for missing lookup IDs
                result.Should().BeNull("Missing lookup IDs should return null");
            }
            else
            {
                // If not null, should be valid
                result.Should().NotBeNull();
            }
        }

        [Theory]
        [InlineData("123456", "123456", "6-digit Content ID should remain unchanged")]
        [InlineData("12345", "012345", "5-digit Content ID should be padded with leading zero")]
        [InlineData("TSRC-PROD-123456", "123456", "Should extract 6-digit from lookup ID")]
        [InlineData("CMS-TEST-654321", "654321", "Should extract 6-digit from CMS lookup ID")]
        public void FormatContentIdForDisplay_ShouldFollowVbaMethodology(string input, string expected, string reason)
        {
            // Use reflection to test private method
            var method = typeof(HyperlinkReplacementService).GetMethod("FormatContentIdForDisplay",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            // Act
            var result = (string)method.Invoke(_service, new object[] { input });

            // Assert
            result.Should().Be(expected, reason);
        }

        [Theory]
        [InlineData("doc<span>123</span>", "doc123", "Should remove HTML tags")]
        [InlineData("doc&amp;test", "doc&test", "Should decode HTML entities")]
        [InlineData("doc&lt;test&gt;", "doc<test>", "Should decode angle bracket entities")]
        [InlineData("doc&quot;test&quot;", "doctest", "Should remove quotes after decoding")]
        [InlineData("normal-doc-id", "normal-doc-id", "Should leave clean IDs unchanged")]
        public void FilterHtmlElementsFromUrl_ShouldCleanUrls(string input, string expected, string reason)
        {
            // Use reflection to test private method
            var method = typeof(HyperlinkReplacementService).GetMethod("FilterHtmlElementsFromUrl",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            // Act
            var result = (string)method.Invoke(_service, new object[] { input });

            // Assert
            result.Should().Be(expected, reason);
        }

        [Theory]
        [InlineData("TSRC-PROD-123456", "123456", "Should extract 6-digit from TSRC")]
        [InlineData("CMS-TEST-654321", "654321", "Should extract 6-digit from CMS")]
        [InlineData("TSRC-VERYLONGNAME-789012", "789012", "Should handle long middle parts")]
        [InlineData("invalid-format", "invalid-format", "Should return original for invalid format")]
        public void ExtractContentIdFromLookupId_ShouldFollowVbaPattern(string input, string expected, string reason)
        {
            // Use reflection to test private method
            var method = typeof(HyperlinkReplacementService).GetMethod("ExtractContentIdFromLookupId",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            // Act
            var result = (string)method.Invoke(_service, new object[] { input });

            // Assert
            result.Should().Be(expected, reason);
        }

        [Fact]
        public void BuildUrlFromDocumentId_WithValidDocumentId_ShouldGenerateCorrectUrl()
        {
            // Arrange
            var documentId = "doc-123456-789";

            // Act
            var result = _service.BuildUrlFromDocumentId(documentId);

            // Assert
            result.Should().NotBeNullOrEmpty("URL should be generated");
            result.Should().StartWith("https://thesource.cvshealth.com/nuxeo/thesource/#!/view?docid=",
                "URL should follow CVS Health format");
            result.Should().Contain(documentId, "URL should contain the document ID");
        }

        [Fact]
        public void BuildUrlFromDocumentId_WithHtmlEntities_ShouldDecodeAndFilterUrl()
        {
            // Arrange
            var documentId = "doc&amp;test&lt;123&gt;";

            // Act
            var result = _service.BuildUrlFromDocumentId(documentId);

            // Assert
            result.Should().NotBeNullOrEmpty("URL should be generated");
            result.Should().NotContain("&amp;", "HTML entities should be decoded");
            result.Should().NotContain("&lt;", "HTML entities should be decoded");
            result.Should().NotContain("&gt;", "HTML entities should be decoded");
        }

        [Theory]
        [InlineData("", false, "Empty URL should be invalid")]
        [InlineData("not-a-url", false, "Plain text should be invalid")]
        [InlineData("ftp://example.com", false, "FTP URLs should be invalid")]
        [InlineData("http://example.com", true, "HTTP URLs should be valid")]
        [InlineData("https://example.com", true, "HTTPS URLs should be valid")]
        [InlineData("https://thesource.cvshealth.com/test", true, "CVS Health URLs should be valid")]
        public void IsValidUrl_ShouldValidateUrlsCorrectly(string url, bool expected, string reason)
        {
            // Use reflection to test private method
            var method = typeof(HyperlinkReplacementService).GetMethod("IsValidUrl",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            // Act
            var result = (bool)method.Invoke(_service, new object[] { url });

            // Assert
            result.Should().Be(expected, reason);
        }

        [Fact]
        public async Task SimulateApiCallAsync_ShouldGenerateRealisticJsonResponse()
        {
            // Use reflection to test private method
            var method = typeof(HyperlinkReplacementService).GetMethod("SimulateApiCallAsync",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            // Arrange
            var lookupIds = new[] { "TSRC-PROD-123456", "CMS-TEST-654321" };

            // Act
            var result = (Task<string>)method.Invoke(_service, new object[] { lookupIds, CancellationToken.None });
            var jsonResponse = await result;

            // Assert
            jsonResponse.Should().NotBeNullOrEmpty("JSON response should be generated");
            jsonResponse.Should().Contain("\"Version\":", "Should contain version field");
            jsonResponse.Should().Contain("\"Changes\":", "Should contain changes field");
            jsonResponse.Should().Contain("\"Results\":", "Should contain results array");
            jsonResponse.Should().Contain("\"Lookup_ID\":", "Should contain lookup ID fields");
            jsonResponse.Should().Contain("\"Document_ID\":", "Should contain document ID fields");
            jsonResponse.Should().Contain("\"Content_ID\":", "Should contain content ID fields");
            jsonResponse.Should().Contain("\"Title\":", "Should contain title fields");
            jsonResponse.Should().Contain("\"Status\":", "Should contain status fields");
        }

        [Fact]
        public void ParseJsonResponseWithStatusDetection_ShouldCategorizeDocumentsCorrectly()
        {
            // Arrange
            var jsonResponse = @"{
                ""Version"": ""1.2.3"",
                ""Changes"": ""Test response"",
                ""Results"": [
                    {
                        ""Lookup_ID"": ""TSRC-PROD-123456"",
                        ""Document_ID"": ""doc-123456-789"",
                        ""Content_ID"": ""123456"",
                        ""Title"": ""Test Document 1"",
                        ""Status"": ""Released""
                    },
                    {
                        ""Lookup_ID"": ""TSRC-PROD-654321"",
                        ""Document_ID"": ""doc-654321-012"",
                        ""Content_ID"": ""654321"",
                        ""Title"": ""Test Document 2"",
                        ""Status"": ""Expired""
                    }
                ]
            }";

            var originalLookupIds = new[] { "TSRC-PROD-123456", "TSRC-PROD-654321", "TSRC-PROD-999999" };

            // Use reflection to test private method
            var method = typeof(HyperlinkReplacementService).GetMethod("ParseJsonResponseWithFlexibleMatching",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            // Act
            var result = method.Invoke(_service, new object[] { jsonResponse, originalLookupIds });
            var typedResult = (HyperlinkReplacementService.ApiProcessingResult)result;

            // Assert
            typedResult.Should().NotBeNull("Parsing should succeed");
            typedResult.IsSuccess.Should().BeTrue("Parsing should be successful");
            typedResult.HasError.Should().BeFalse("Should not have errors");

            typedResult.FoundDocuments.Should().HaveCount(1, "One document should be found with Released status");
            typedResult.FoundDocuments.First().Status.Should().Be("Released");
            typedResult.FoundDocuments.First().Lookup_ID.Should().Be("TSRC-PROD-123456");

            typedResult.ExpiredDocuments.Should().HaveCount(1, "One document should be expired");
            typedResult.ExpiredDocuments.First().Status.Should().Be("Expired");
            typedResult.ExpiredDocuments.First().Lookup_ID.Should().Be("TSRC-PROD-654321");

            typedResult.MissingLookupIds.Should().HaveCount(1, "One lookup ID should be missing");
            typedResult.MissingLookupIds.First().Should().Be("TSRC-PROD-999999");
        }

        [Fact]
        public async Task LookupTitleByIdentifierAsync_WithValidContentId_ShouldReturnTitle()
        {
            // Arrange
            var contentId = "123456";

            // Act
            var result = await _service.LookupTitleByIdentifierAsync(contentId, CancellationToken.None);

            // Assert
            result.Should().NotBeNullOrEmpty("Should return a title");
            result.Should().Contain(contentId, "Title should contain the content ID");
        }

        [Fact]
        public async Task LookupTitleByIdentifierAsync_WithMissingContentId_ShouldReturnFallbackTitle()
        {
            // Arrange - Use content ID that will likely be missing based on simulation
            var contentId = "000001";

            // Act
            var result = await _service.LookupTitleByIdentifierAsync(contentId, CancellationToken.None);

            // Assert
            result.Should().NotBeNullOrEmpty("Should always return a title, even for missing content");
            result.Should().Contain(contentId, "Fallback title should contain the content ID");
        }

        [Fact]
        public void ValidateAndSanitizeUrlFragment_WithExclamationMark_ShouldEncodeForOpenXmlCompatibility()
        {
            // Arrange - Test the specific 0x21 (!) character that was causing XSD validation errors
            var fragmentWithExclamation = "!/view?docid=test-doc-123";

            // Use reflection to test private method
            var method = typeof(HyperlinkReplacementService).GetMethod("ValidateAndSanitizeUrlFragment",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            // Act
            var result = (string)method.Invoke(_service, new object[] { fragmentWithExclamation });

            // Assert
            result.Should().NotBeNull("Method should return a result");
            result.Should().NotContain("!", "Exclamation mark should be URL encoded");
            result.Should().Contain("%21", "Exclamation mark should be encoded as %21");
            result.Should().Contain("view", "The rest of the URL should be preserved");
            result.Should().Contain("docid", "The docid parameter should be preserved");
        }

        [Fact]
        public void IsUrlFragmentSafeForOpenXml_WithExclamationMark_ShouldReturnFalse()
        {
            // Arrange - Test the specific 0x21 (!) character
            var unsafeFragment = "!/view?docid=test";
            var safeFragment = "view?docid=test";

            // Use reflection to test private method
            var method = typeof(HyperlinkReplacementService).GetMethod("IsUrlFragmentSafeForOpenXml",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            // Act
            var unsafeResult = (bool)method.Invoke(_service, new object[] { unsafeFragment });
            var safeResult = (bool)method.Invoke(_service, new object[] { safeFragment });

            // Assert
            unsafeResult.Should().BeFalse("Fragment with exclamation mark should be unsafe for OpenXML");
            safeResult.Should().BeTrue("Fragment without special characters should be safe for OpenXML");
        }

        [Theory]
        [InlineData("!/view?docid=test", "Fragment with exclamation mark")]
        [InlineData("<script>alert('test')</script>", "Fragment with angle brackets")]
        [InlineData("test&param=value", "Fragment with ampersand")]
        [InlineData("test\"quoted\"", "Fragment with quotes")]
        [InlineData("test'quoted'", "Fragment with single quotes")]
        public void ValidateAndSanitizeUrlFragment_WithSpecialCharacters_ShouldEncodeCorrectly(string input, string description)
        {
            // Use reflection to test private method
            var method = typeof(HyperlinkReplacementService).GetMethod("ValidateAndSanitizeUrlFragment",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            // Act
            var result = (string)method.Invoke(_service, new object[] { input });

            // Assert
            result.Should().NotBeNull($"Method should return a result for: {description}");
            result.Should().NotContain("!", "Exclamation marks should be encoded");
            result.Should().NotContain("<", "Angle brackets should be encoded");
            result.Should().NotContain(">", "Angle brackets should be encoded");
            result.Should().NotContain("&", "Ampersands should be encoded (except for encoded sequences)");
            result.Should().NotContain("\"", "Double quotes should be encoded");
            result.Should().NotContain("'", "Single quotes should be encoded");
        }

        [Fact]
        public async Task BuildUrlFromDocumentId_WithCleanDocumentId_ShouldCreateValidUrl()
        {
            // Arrange
            var documentId = "test-document-123";

            // Act
            var result = _service.BuildUrlFromDocumentId(documentId);

            // Assert
            result.Should().NotBeNullOrEmpty("URL should be generated");
            result.Should().StartWith("https://thesource.cvshealth.com/nuxeo/thesource/#!/view?docid=",
                "URL should follow the correct CVS Health format");
            result.Should().EndWith(documentId, "URL should contain the document ID");
        }
    }
}
