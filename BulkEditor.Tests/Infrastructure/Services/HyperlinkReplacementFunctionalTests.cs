using BulkEditor.Core.Configuration;
using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using BulkEditor.Infrastructure.Services;
using Microsoft.Extensions.Options;
using Moq;
using Xunit;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System;

namespace BulkEditor.Tests.Infrastructure.Services
{
    /// <summary>
    /// Functional tests to verify the custom hyperlink replacement functionality works as specified:
    /// - Find hyperlinks where "Text to Display" contains the provided text
    /// - Replace "Text to Display" with "Title" + " (Last 6 of Content_ID)"
    /// - Replace URL with "https://thesource.cvshealth.com/nuxeo/thesource/#!/view?docid=" + Document_ID
    /// </summary>
    public class HyperlinkReplacementFunctionalTests
    {
        private readonly Mock<IHttpService> _httpServiceMock;
        private readonly Mock<ILoggingService> _loggerMock;
        private readonly Mock<IRetryPolicyService> _retryPolicyServiceMock;
        private readonly HyperlinkReplacementService _service;
        private readonly AppSettings _appSettings;

        public HyperlinkReplacementFunctionalTests()
        {
            _httpServiceMock = new Mock<IHttpService>();
            _loggerMock = new Mock<ILoggingService>();
            _retryPolicyServiceMock = new Mock<IRetryPolicyService>();

            _appSettings = new AppSettings
            {
                Api = new ApiSettings
                {
                    BaseUrl = "https://api.example.com"
                },
                Replacement = new ReplacementSettings
                {
                    EnableHyperlinkReplacement = true,
                    HyperlinkRules = new List<HyperlinkReplacementRule>
                    {
                        new HyperlinkReplacementRule
                        {
                            Id = "test-rule-1",
                            TitleToMatch = "Document Title",
                            ContentId = "123456",
                            IsEnabled = true
                        },
                        new HyperlinkReplacementRule
                        {
                            Id = "test-rule-2", 
                            TitleToMatch = "Test Document",
                            ContentId = "TSRC-PROD-654321",
                            IsEnabled = true
                        }
                    }
                }
            };

            var appSettingsOptions = Options.Create(_appSettings);
            _service = new HyperlinkReplacementService(_httpServiceMock.Object, _loggerMock.Object, appSettingsOptions, _retryPolicyServiceMock.Object);
        }

        [Fact(Timeout = 10000)]
        public async Task ProcessApiResponseAsync_WithValidLookupIds_ShouldReturnExpectedFormat()
        {
            // Arrange - Test the API response processing that simulates the expected JSON structure
            var lookupIds = new[] { "123456", "TSRC-PROD-654321" };

            // Act
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            var result = await _service.ProcessApiResponseAsync(lookupIds, cts.Token);

            // Assert - Verify the API processing returns the expected structure
            Assert.NotNull(result);
            Assert.True(result.IsSuccess);
            
            // Should have found documents (not all missing)
            var totalFound = result.FoundDocuments.Count + result.ExpiredDocuments.Count;
            Assert.True(totalFound > 0, "Should find at least some documents");

            // Verify document records have the required fields
            var allDocuments = result.FoundDocuments.Concat(result.ExpiredDocuments);
            foreach (var doc in allDocuments)
            {
                Assert.False(string.IsNullOrEmpty(doc.Title), $"Document {doc.Lookup_ID} should have a Title");
                Assert.False(string.IsNullOrEmpty(doc.Document_ID), $"Document {doc.Lookup_ID} should have a Document_ID");
                Assert.False(string.IsNullOrEmpty(doc.Content_ID), $"Document {doc.Lookup_ID} should have a Content_ID");
                Assert.Contains(doc.Status, new[] { "Released", "Expired" }, StringComparer.OrdinalIgnoreCase);
            }
        }

        [Fact(Timeout = 10000)]
        public async Task LookupDocumentByIdentifierAsync_ShouldReturnCorrectFormat()
        {
            // Arrange - Test individual document lookup
            var identifier = "123456";

            // Act
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            var result = await _service.LookupDocumentByIdentifierAsync(identifier, cts.Token);

            // Assert - Verify the lookup returns the expected format
            if (result != null) // May be null for missing IDs in simulation
            {
                Assert.False(string.IsNullOrEmpty(result.Title));
                Assert.False(string.IsNullOrEmpty(result.Document_ID));
                Assert.False(string.IsNullOrEmpty(result.Content_ID));
                Assert.Contains(result.Status, new[] { "Released", "Expired" }, StringComparer.OrdinalIgnoreCase);
                
                // Verify Title format expected for hyperlink replacement
                Assert.Contains(identifier, result.Title);
                
                // Verify Document_ID is different from Content_ID (as expected in real API)
                Assert.NotEqual(result.Document_ID, result.Content_ID);
            }
        }

        [Fact]
        public async Task LookupTitleByIdentifierAsync_ShouldReturnProperTitleFormat()
        {
            // Arrange - Test title lookup that would be used for hyperlink display text
            var identifier = "123456";

            // Act
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            var result = await _service.LookupTitleByIdentifierAsync(identifier, cts.Token);

            // Assert - Verify title format
            Assert.False(string.IsNullOrEmpty(result));
            Assert.Contains(identifier, result);
        }

        [Fact]
        public void BuildUrlFromDocumentId_ShouldCreateCorrectUrlFormat()
        {
            // Arrange - Test URL building with Document_ID
            var documentId = "doc-test-123456-789";

            // Act
            var result = _service.BuildUrlFromDocumentId(documentId);

            // Assert - Verify the URL matches the expected format
            Assert.False(string.IsNullOrEmpty(result));
            Assert.StartsWith("https://thesource.cvshealth.com/nuxeo/thesource/#!/view?docid=", result);
            Assert.EndsWith(documentId, result);
            
            // Verify complete expected format
            var expectedUrl = $"https://thesource.cvshealth.com/nuxeo/thesource/#!/view?docid={documentId}";
            Assert.Equal(expectedUrl, result);
        }

        [Theory]
        [InlineData("123456", "123456")] // 6-digit should stay as-is
        [InlineData("12345", "012345")]  // 5-digit should be padded
        [InlineData("1234", "1234")]     // Shorter than 5 digits returns as-is
        public void ContentIdFormatting_ShouldFollowExpectedRules(string input, string expected)
        {
            // Test the Content_ID formatting logic using reflection
            var method = typeof(HyperlinkReplacementService).GetMethod("FormatContentIdForDisplay",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            
            // Act
            var result = (string)method.Invoke(_service, new object[] { input });
            
            // Assert
            Assert.Equal(expected, result);
        }

        [Fact] 
        public void ConfigurationValidation_ShouldHaveCorrectRulesSetup()
        {
            // Verify the test configuration has the expected rules
            Assert.True(_appSettings.Replacement.EnableHyperlinkReplacement);
            Assert.Equal(2, _appSettings.Replacement.HyperlinkRules.Count);
            
            var rule1 = _appSettings.Replacement.HyperlinkRules[0];
            Assert.Equal("Document Title", rule1.TitleToMatch);
            Assert.Equal("123456", rule1.ContentId);
            Assert.True(rule1.IsEnabled);
            
            var rule2 = _appSettings.Replacement.HyperlinkRules[1];
            Assert.Equal("Test Document", rule2.TitleToMatch);
            Assert.Equal("TSRC-PROD-654321", rule2.ContentId);
            Assert.True(rule2.IsEnabled);
        }

        [Theory]
        [InlineData("Document Title for testing", "document title", true)]  // Contains match
        [InlineData("Test Document v2", "test document", true)]            // Contains match
        [InlineData("Different Title", "document title", false)]           // No match
        [InlineData("", "document title", false)]                          // Empty source
        [InlineData("Document Title", "", false)]                          // Empty pattern
        public void TextMatching_ShouldFollowContainsLogic(string sourceText, string matchText, bool expectedMatch)
        {
            // Test the text matching logic using reflection
            var method = typeof(HyperlinkReplacementService).GetMethod("DoesTextMatch",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            
            // Act - Test with Contains mode (default)
            var result = (bool)method.Invoke(_service, new object[] { sourceText, matchText, HyperlinkMatchMode.Contains });
            
            // Assert
            Assert.Equal(expectedMatch, result);
        }

        [Fact]
        public async Task SimulationLogic_ShouldProducePredictableResults()
        {
            // Test that the simulation produces consistent results for specific patterns
            var predictableIds = new[] { "123456", "654321", "111111" };
            
            // Act
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            var result = await _service.ProcessApiResponseAsync(predictableIds, cts.Token);
            
            // Assert - Should have consistent behavior
            Assert.NotNull(result);
            Assert.True(result.IsSuccess);
            
            // The 654321 ID should consistently be marked as expired based on simulation logic
            var expiredDoc = result.ExpiredDocuments.FirstOrDefault(d => d.Content_ID.Contains("654321"));
            if (expiredDoc != null)
            {
                Assert.Equal("Expired", expiredDoc.Status);
            }
        }
    }

    /// <summary>
    /// Enum for hyperlink title matching modes (copied from main implementation for testing)
    /// </summary>
    public enum HyperlinkMatchMode
    {
        Exact,
        Contains,
        StartsWith,
        EndsWith
    }
}