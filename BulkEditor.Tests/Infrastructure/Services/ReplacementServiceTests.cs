using BulkEditor.Core.Configuration;
using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using BulkEditor.Infrastructure.Services;
using FluentAssertions;
using Moq;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

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
        public async Task ProcessReplacementsAsync_WhenNoReplacementsEnabled_ShouldNotCallServices()
        {
            // Arrange
            var document = new Document { FileName = "test.docx" };

            // Act
            var result = await _service.ProcessReplacementsAsync(document, CancellationToken.None);

            // Assert
            result.Should().BeSameAs(document);
            _mockHyperlinkService.Verify(x => x.ProcessHyperlinkReplacementsAsync(It.IsAny<Document>(), It.IsAny<IEnumerable<HyperlinkReplacementRule>>(), It.IsAny<CancellationToken>()), Times.Never);
            _mockTextService.Verify(x => x.ProcessTextReplacementsAsync(It.IsAny<Document>(), It.IsAny<IEnumerable<TextReplacementRule>>(), It.IsAny<CancellationToken>()), Times.Never);
        }

        [Fact]
        public async Task ProcessReplacementsAsync_WhenHyperlinkReplacementEnabled_ShouldCallHyperlinkService()
        {
            // Arrange
            var document = new Document { FileName = "test.docx" };
            _appSettings.Replacement.EnableHyperlinkReplacement = true;
            _appSettings.Replacement.HyperlinkRules.Add(new HyperlinkReplacementRule
            {
                TitleToMatch = "Test Title",
                ContentId = "123456"
            });

            _mockHyperlinkService.Setup(x => x.ProcessHyperlinkReplacementsAsync(It.IsAny<Document>(), It.IsAny<IEnumerable<HyperlinkReplacementRule>>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(document);

            // Act
            var result = await _service.ProcessReplacementsAsync(document, CancellationToken.None);

            // Assert
            result.Should().BeSameAs(document);
            _mockHyperlinkService.Verify(x => x.ProcessHyperlinkReplacementsAsync(document, It.IsAny<IEnumerable<HyperlinkReplacementRule>>(), It.IsAny<CancellationToken>()), Times.Once);
        }

        [Fact]
        public async Task ProcessReplacementsAsync_WhenTextReplacementEnabled_ShouldCallTextService()
        {
            // Arrange
            var document = new Document { FileName = "test.docx" };
            _appSettings.Replacement.EnableTextReplacement = true;
            _appSettings.Replacement.TextRules.Add(new TextReplacementRule
            {
                SourceText = "old text",
                ReplacementText = "new text"
            });

            _mockTextService.Setup(x => x.ProcessTextReplacementsAsync(It.IsAny<Document>(), It.IsAny<IEnumerable<TextReplacementRule>>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(document);

            // Act
            var result = await _service.ProcessReplacementsAsync(document, CancellationToken.None);

            // Assert
            result.Should().BeSameAs(document);
            _mockTextService.Verify(x => x.ProcessTextReplacementsAsync(document, It.IsAny<IEnumerable<TextReplacementRule>>(), It.IsAny<CancellationToken>()), Times.Once);
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
            result.IsValid.Should().BeTrue();
            result.ValidRulesCount.Should().Be(1);
            result.InvalidRulesCount.Should().Be(0);
            result.ValidationErrors.Should().BeEmpty();
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
            result.IsValid.Should().BeFalse();
            result.ValidRulesCount.Should().Be(0);
            result.InvalidRulesCount.Should().Be(1);
            result.ValidationErrors.Should().HaveCount(2);
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
            result.IsValid.Should().BeTrue();
            result.ValidRulesCount.Should().Be(1);
            result.InvalidRulesCount.Should().Be(0);
            result.ValidationErrors.Should().BeEmpty();
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
            result.IsValid.Should().BeFalse();
            result.InvalidRulesCount.Should().Be(1);
            result.ValidationErrors.Should().ContainSingle();
        }
    }
}