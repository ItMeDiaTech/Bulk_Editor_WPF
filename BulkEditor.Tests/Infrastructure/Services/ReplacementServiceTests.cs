using BulkEditor.Core.Configuration;
using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using BulkEditor.Infrastructure.Services;
using System.Linq;
using Moq;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Xunit;
using DocumentFormat.OpenXml.Packaging;

namespace BulkEditor.Tests.Infrastructure.Services
{
    public class ReplacementServiceTests
    {
        private readonly Mock<IHyperlinkReplacementService> _mockHyperlinkService;
        private readonly Mock<ITextReplacementService> _mockTextService;
        private readonly Mock<ILoggingService> _mockLogger;
        private readonly AppSettings _appSettings;
        private readonly ReplacementService _service;

        public ReplacementServiceTests()
        {
            _mockHyperlinkService = new Mock<IHyperlinkReplacementService>();
            _mockTextService = new Mock<ITextReplacementService>();
            _mockLogger = new Mock<ILoggingService>();

            _appSettings = new AppSettings
            {
                Replacement = new ReplacementSettings
                {
                    EnableHyperlinkReplacement = false,
                    EnableTextReplacement = false,
                    HyperlinkRules = new List<HyperlinkReplacementRule>(),
                    TextRules = new List<TextReplacementRule>()
                }
            };

            _service = new ReplacementService(
                _mockHyperlinkService.Object,
                _mockTextService.Object,
                _mockLogger.Object,
                _appSettings);
        }

        [Fact]
        public async Task ProcessReplacementsInSessionAsync_WhenNoReplacementsEnabled_ShouldNotCallServices()
        {
            // Arrange
            var document = new Document { FileName = "test.docx" };
            // Note: Using null for WordprocessingDocument since we're testing coordination logic, not OpenXML operations

            // Act
            var result = await _service.ProcessReplacementsInSessionAsync(null, document, CancellationToken.None);

            // Assert
            Assert.Equal(0, result);
            _mockHyperlinkService.Verify(x => x.ProcessHyperlinkReplacementsInSessionAsync(It.IsAny<WordprocessingDocument>(), It.IsAny<Document>(), It.IsAny<IEnumerable<HyperlinkReplacementRule>>(), It.IsAny<CancellationToken>()), Times.Never);
            _mockTextService.Verify(x => x.ProcessTextReplacementsInSessionAsync(It.IsAny<WordprocessingDocument>(), It.IsAny<Document>(), It.IsAny<IEnumerable<TextReplacementRule>>(), It.IsAny<CancellationToken>()), Times.Never);
        }

        [Fact]
        public async Task ProcessReplacementsInSessionAsync_WhenHyperlinkReplacementEnabled_ShouldCallHyperlinkService()
        {
            // Arrange
            var document = new Document { FileName = "test.docx" };
            _appSettings.Replacement.EnableHyperlinkReplacement = true;
            _appSettings.Replacement.HyperlinkRules.Add(new HyperlinkReplacementRule
            {
                TitleToMatch = "Test Title",
                ContentId = "123456"
            });

            _mockHyperlinkService.Setup(x => x.ProcessHyperlinkReplacementsInSessionAsync(It.IsAny<WordprocessingDocument>(), It.IsAny<Document>(), It.IsAny<IEnumerable<HyperlinkReplacementRule>>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(1);
            // Note: Using null for WordprocessingDocument since we're testing coordination logic, not OpenXML operations

            // Act
            var result = await _service.ProcessReplacementsInSessionAsync(null, document, CancellationToken.None);

            // Assert
            Assert.Equal(1, result);
            _mockHyperlinkService.Verify(x => x.ProcessHyperlinkReplacementsInSessionAsync(It.IsAny<WordprocessingDocument>(), document, It.IsAny<IEnumerable<HyperlinkReplacementRule>>(), It.IsAny<CancellationToken>()), Times.Once);
        }

        [Fact]
        public async Task ProcessReplacementsInSessionAsync_WhenTextReplacementEnabled_ShouldCallTextService()
        {
            // Arrange
            var document = new Document { FileName = "test.docx" };
            _appSettings.Replacement.EnableTextReplacement = true;
            _appSettings.Replacement.TextRules.Add(new TextReplacementRule
            {
                SourceText = "old text",
                ReplacementText = "new text"
            });

            _mockTextService.Setup(x => x.ProcessTextReplacementsInSessionAsync(It.IsAny<WordprocessingDocument>(), It.IsAny<Document>(), It.IsAny<IEnumerable<TextReplacementRule>>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(1);
            // Note: Using null for WordprocessingDocument since we're testing coordination logic, not OpenXML operations

            // Act
            var result = await _service.ProcessReplacementsInSessionAsync(null, document, CancellationToken.None);

            // Assert
            Assert.Equal(1, result);
            _mockTextService.Verify(x => x.ProcessTextReplacementsInSessionAsync(It.IsAny<WordprocessingDocument>(), document, It.IsAny<IEnumerable<TextReplacementRule>>(), It.IsAny<CancellationToken>()), Times.Once);
        }

        [Fact]
        public async Task ValidateReplacementRulesAsync_WithValidHyperlinkRule_ShouldReturnValid()
        {
            // Arrange
            var rules = new List<object>
            {
                new HyperlinkReplacementRule
                {
                    TitleToMatch = "Test Title",
                    ContentId = "123456"
                }
            };

            // Act
            var result = await _service.ValidateReplacementRulesAsync(rules, CancellationToken.None);

            // Assert
            Assert.True(result.IsValid);
            Assert.Equal(1, result.ValidRulesCount);
            Assert.Equal(0, result.InvalidRulesCount);
            Assert.Empty(result.ValidationErrors);
        }

        [Fact]
        public async Task ValidateReplacementRulesAsync_WithInvalidHyperlinkRule_ShouldReturnInvalid()
        {
            // Arrange
            var rules = new List<object>
            {
                new HyperlinkReplacementRule
                {
                    TitleToMatch = "", // Invalid: empty title
                    ContentId = "abc" // Invalid: not 6 digits
                }
            };

            // Act
            var result = await _service.ValidateReplacementRulesAsync(rules, CancellationToken.None);

            // Assert
            Assert.False(result.IsValid);
            Assert.Equal(0, result.ValidRulesCount);
            Assert.Equal(1, result.InvalidRulesCount);
            Assert.Equal(2, result.ValidationErrors.Count);
        }

        [Fact]
        public async Task ValidateReplacementRulesAsync_WithValidTextRule_ShouldReturnValid()
        {
            // Arrange
            var rules = new List<object>
            {
                new TextReplacementRule
                {
                    SourceText = "old text",
                    ReplacementText = "new text"
                }
            };

            // Act
            var result = await _service.ValidateReplacementRulesAsync(rules, CancellationToken.None);

            // Assert
            Assert.True(result.IsValid);
            Assert.Equal(1, result.ValidRulesCount);
            Assert.Equal(0, result.InvalidRulesCount);
            Assert.Empty(result.ValidationErrors);
        }

        [Fact]
        public async Task ValidateReplacementRulesAsync_WithSameSourceAndReplacement_ShouldReturnInvalid()
        {
            // Arrange
            var rules = new List<object>
            {
                new TextReplacementRule
                {
                    SourceText = "same text",
                    ReplacementText = "same text" // Invalid: same as source
                }
            };

            // Act
            var result = await _service.ValidateReplacementRulesAsync(rules, CancellationToken.None);

            // Assert
            Assert.False(result.IsValid);
            Assert.Equal(1, result.InvalidRulesCount);
            Assert.Single(result.ValidationErrors);
        }
    }
}