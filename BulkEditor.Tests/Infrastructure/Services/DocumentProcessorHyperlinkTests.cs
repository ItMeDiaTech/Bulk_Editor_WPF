using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using BulkEditor.Infrastructure.Services;
using Moq;
using Xunit;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;

namespace BulkEditor.Tests.Infrastructure.Services
{
    /// <summary>
    /// Tests for hyperlink filtering logic to ensure VBA compatibility
    /// These tests verify that C# implementation matches Base_File.vba exactly
    /// </summary>
    public class DocumentProcessorHyperlinkTests
    {
        private readonly Mock<IFileService> _fileServiceMock;
        private readonly Mock<IHyperlinkValidator> _hyperlinkValidatorMock;
        private readonly Mock<ITextOptimizer> _textOptimizerMock;
        private readonly Mock<IReplacementService> _replacementServiceMock;
        private readonly Mock<ILoggingService> _loggerMock;
        private readonly BulkEditor.Core.Configuration.AppSettings _appSettings;
        private readonly DocumentProcessor _documentProcessor;

        public DocumentProcessorHyperlinkTests()
        {
            _fileServiceMock = new Mock<IFileService>();
            _hyperlinkValidatorMock = new Mock<IHyperlinkValidator>();
            _textOptimizerMock = new Mock<ITextOptimizer>();
            _replacementServiceMock = new Mock<IReplacementService>();
            _loggerMock = new Mock<ILoggingService>();

            _appSettings = new BulkEditor.Core.Configuration.AppSettings
            {
                Processing = new BulkEditor.Core.Configuration.ProcessingSettings
                {
                    MaxConcurrentDocuments = 2
                },
                Validation = new BulkEditor.Core.Configuration.ValidationSettings
                {
                    AutoReplaceTitles = false,
                    ReportTitleDifferences = true
                }
            };

            var mockRetryPolicyService = new Mock<BulkEditor.Core.Services.IRetryPolicyService>();
            
            _documentProcessor = new DocumentProcessor(
                _fileServiceMock.Object,
                _hyperlinkValidatorMock.Object,
                _textOptimizerMock.Object,
                _replacementServiceMock.Object,
                _loggerMock.Object,
                _appSettings,
                mockRetryPolicyService.Object);
        }

        [Theory]
        [InlineData("https://example.com/page?docid=TSRC-PROD-123456", true, "Should validate TSRC with exact pattern")]
        [InlineData("https://example.com/page?docid=CMS-PRD1-654321", true, "Should validate CMS with exact pattern")]
        [InlineData("https://example.com/page?docid=TSRC-TEST-123456", true, "Should validate TSRC-TEST format")]
        [InlineData("https://example.com/page?docid=CMS-DEV-654321", true, "Should validate CMS-DEV format")]
        [InlineData("https://thesource.cvshealth.com/nuxeo/thesource/#!/view?docid=doc-411364-329", true, "Should validate docid parameter")]
        [InlineData("https://example.com/TSRC-PROD-123456/page", true, "Should validate TSRC in URL path")]
        [InlineData("https://example.com/CMS-PRD1-654321/page", true, "Should validate CMS in URL path")]
        public void ShouldAutoValidateHyperlink_ValidFormats_ReturnsTrue(string url, bool expected, string reason)
        {
            // Arrange
            _hyperlinkValidatorMock.Setup(x => x.ExtractLookupId(It.IsAny<string>()))
                .Returns("EXTRACTED_ID"); // Mock returns valid ID for valid URLs

            // Act
            var result = InvokeShouldAutoValidateHyperlink(url, "Test Display Text");

            // Assert
            Assert.Equal(expected, result);
        }

        [Theory]
        [InlineData("https://example.com/TSRC123456", false, "Should NOT validate TSRC without proper separators")]
        [InlineData("https://example.com/CMS123456", false, "Should NOT validate CMS without proper separators")]
        [InlineData("https://example.com/TSRC-", false, "Should NOT validate incomplete TSRC pattern")]
        [InlineData("https://example.com/CMS-", false, "Should NOT validate incomplete CMS pattern")]
        [InlineData("https://example.com/TSRC-PROD-12345", false, "Should NOT validate TSRC with 5 digits instead of 6")]
        [InlineData("https://example.com/CMS-PRD1-1234567", false, "Should NOT validate CMS with 7 digits instead of 6")]
        [InlineData("https://example.com/TSRCx-PROD-123456", false, "Should NOT validate TSRC with extra characters")]
        [InlineData("https://example.com/CMSx-PRD1-123456", false, "Should NOT validate CMS with extra characters")]
        [InlineData("https://example.com/some-tsrc-page", false, "Should NOT validate loose TSRC text")]
        [InlineData("https://example.com/some-cms-page", false, "Should NOT validate loose CMS text")]
        [InlineData("https://www.google.com", false, "Should NOT validate random URLs")]
        [InlineData("", false, "Should NOT validate empty URLs")]
        public void ShouldAutoValidateHyperlink_InvalidFormats_ReturnsFalse(string url, bool expected, string reason)
        {
            // Arrange
            _hyperlinkValidatorMock.Setup(x => x.ExtractLookupId(It.IsAny<string>()))
                .Returns(""); // Mock returns empty for invalid URLs

            // Act
            var result = InvokeShouldAutoValidateHyperlink(url, "Test Display Text");

            // Assert
            Assert.Equal(expected, result);
        }

        [Theory]
        [InlineData("TSRC-PROD-123456", "TSRC-PROD-123456", "Should extract exact TSRC pattern")]
        [InlineData("CMS-PRD1-654321", "CMS-PRD1-654321", "Should extract exact CMS pattern")]
        [InlineData("tsrc-test-123456", "TSRC-TEST-123456", "Should extract and uppercase TSRC")]
        [InlineData("cms-dev-654321", "CMS-DEV-654321", "Should extract and uppercase CMS")]
        [InlineData("TSRC-PROD-123456-suffix", "TSRC-PROD-123456", "Should extract TSRC from longer string")]
        [InlineData("prefix-CMS-PRD1-654321-suffix", "CMS-PRD1-654321", "Should extract CMS from longer string")]
        [InlineData("https://example.com/TSRC-PROD-123456/page", "TSRC-PROD-123456", "Should extract TSRC from URL")]
        [InlineData("https://example.com/page?id=CMS-PRD1-654321&other=value", "CMS-PRD1-654321", "Should extract CMS from query params")]
        // User-specific examples that should match VBA pattern
        [InlineData("TSRC-ASDASDASF123132131ASDASD-123456", "TSRC-ASDASDASF123132131ASDASD-123456", "Should extract TSRC with very long middle part")]
        [InlineData("TSRC-PROD1-654321", "TSRC-PROD1-654321", "Should extract TSRC-PROD1 pattern")]
        [InlineData("CMS-21-987654", "CMS-21-987654", "Should extract CMS-21 pattern")]
        [InlineData("https://example.com/TSRC-ASDASDASF123132131ASDASD-123456/page", "TSRC-ASDASDASF123132131ASDASD-123456", "Should extract TSRC with long middle from URL")]
        [InlineData("https://example.com/CMS-21-654321/page", "CMS-21-654321", "Should extract CMS-21 from URL")]
        public void ExtractLookupIdUsingVbaLogic_ValidPatterns_ReturnsCorrectId(string input, string expected, string reason)
        {
            // Act
            var result = InvokeExtractLookupIdUsingVbaLogic(input, "");

            // Assert
            Assert.Equal(expected, result);
        }

        [Theory]
        [InlineData("https://example.com/page?docid=doc-411364-329", "doc-411364-329", "Should extract docid parameter")]
        [InlineData("https://example.com/page?docid=some-document-id&other=value", "some-document-id", "Should extract docid with other params")]
        [InlineData("https://example.com/page?docid=123456", "123456", "Should extract numeric docid")]
        [InlineData("https://example.com/page?other=value&docid=test-doc", "test-doc", "Should extract docid when not first param")]
        public void ExtractLookupIdUsingVbaLogic_DocIdParameter_ReturnsCorrectId(string input, string expected, string reason)
        {
            // Act
            var result = InvokeExtractLookupIdUsingVbaLogic(input, "");

            // Assert
            Assert.Equal(expected, result);
        }

        [Theory]
        [InlineData("TSRC123456", "", "Should NOT extract TSRC without separators")]
        [InlineData("CMS123456", "", "Should NOT extract CMS without separators")]
        [InlineData("TSRC-", "", "Should NOT extract incomplete TSRC")]
        [InlineData("CMS-", "", "Should NOT extract incomplete CMS")]
        [InlineData("TSRC-PROD-12345", "", "Should NOT extract TSRC with wrong digit count")]
        [InlineData("CMS-PRD1-1234567", "", "Should NOT extract CMS with wrong digit count")]
        [InlineData("XSRC-PROD-123456", "", "Should NOT extract wrong prefix")]
        [InlineData("https://www.google.com", "", "Should NOT extract from random URLs")]
        [InlineData("", "", "Should NOT extract from empty string")]
        [InlineData("https://example.com/page?otherparam=value", "", "Should NOT extract when no valid patterns")]
        public void ExtractLookupIdUsingVbaLogic_InvalidPatterns_ReturnsEmpty(string input, string expected, string reason)
        {
            // Act
            var result = InvokeExtractLookupIdUsingVbaLogic(input, "");

            // Assert
            Assert.Equal(expected, result);
        }

        [Theory]
        [InlineData("https://example.com/TSRC-PROD-123456", "additional-info", "https://example.com/TSRC-PROD-123456#additional-info", "Should combine address and subaddress")]
        [InlineData("https://example.com/page", "", "https://example.com/page", "Should handle empty subaddress")]
        [InlineData("https://example.com/CMS-PRD1-654321", "section1", "https://example.com/CMS-PRD1-654321#section1", "Should combine CMS URL with subaddress")]
        public void ExtractLookupIdUsingVbaLogic_SubAddress_CombinesCorrectly(string address, string subAddress, string expectedCombined, string reason)
        {
            // Act - This tests the URL combination logic internally
            var result = InvokeExtractLookupIdUsingVbaLogic(address, subAddress);

            // Assert - We verify behavior by checking if it extracts correctly from combined URL
            if (expectedCombined.Contains("TSRC-") || expectedCombined.Contains("CMS-"))
            {
                Assert.NotEmpty(result);
            }
            else
            {
                Assert.Empty(result);
            }
        }

        [Fact]
        public void HyperlinkFilteringBehavior_MatchesVbaExactly()
        {
            // This test ensures our filtering matches the VBA Base_File.vba logic exactly

            var testCases = new[]
            {
                // VBA WOULD process these (should return true)
                new { Url = "https://example.com/TSRC-PROD-123456", ShouldProcess = true },
                new { Url = "https://example.com/CMS-PRD1-654321", ShouldProcess = true },
                new { Url = "https://example.com/page?docid=doc-123456-789", ShouldProcess = true },

                // VBA would NOT process these (should return false)
                new { Url = "https://example.com/tsrc-content", ShouldProcess = false },
                new { Url = "https://example.com/cms-data", ShouldProcess = false },
                new { Url = "https://example.com/TSRC123456", ShouldProcess = false },
                new { Url = "https://www.google.com", ShouldProcess = false }
            };

            foreach (var testCase in testCases)
            {
                // Setup mock to return appropriate extraction result
                _hyperlinkValidatorMock.Setup(x => x.ExtractLookupId(testCase.Url))
                    .Returns(testCase.ShouldProcess ? "EXTRACTED_ID" : "");

                var result = InvokeShouldAutoValidateHyperlink(testCase.Url, "Test Display");

                Assert.Equal(testCase.ShouldProcess, result);
            }
        }

        /// <summary>
        /// Uses reflection to test private ShouldAutoValidateHyperlink method
        /// </summary>
        private bool InvokeShouldAutoValidateHyperlink(string url, string displayText)
        {
            var method = typeof(DocumentProcessor).GetMethod("ShouldAutoValidateHyperlink",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            return (bool)method.Invoke(_documentProcessor, new object[] { url, displayText });
        }

        /// <summary>
        /// Uses reflection to test private ExtractLookupIdUsingVbaLogic method
        /// </summary>
        private string InvokeExtractLookupIdUsingVbaLogic(string address, string subAddress)
        {
            var method = typeof(DocumentProcessor).GetMethod("ExtractLookupIdUsingVbaLogic",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            return (string)method.Invoke(_documentProcessor, new object[] { address, subAddress });
        }
    }
}