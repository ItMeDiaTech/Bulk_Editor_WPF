using BulkEditor.Core.Interfaces;
using BulkEditor.Infrastructure.Services;
using FluentAssertions;
using Moq;
using Xunit;

namespace BulkEditor.Tests.Infrastructure.Services
{
    public class TextReplacementServiceTests
    {
        private readonly Mock<ILoggingService> _mockLogger;
        private readonly TextReplacementService _service;

        public TextReplacementServiceTests()
        {
            _mockLogger = new Mock<ILoggingService>();
            _service = new TextReplacementService(_mockLogger.Object);
        }

        [Theory]
        [InlineData("Hello world", "world", "universe", "Hello universe")]
        [InlineData("Test text here", "text", "content", "Test content here")]
        [InlineData("Multiple words here", "words", "items", "Multiple items here")]
        public void ReplaceTextWithCapitalizationPreservation_BasicReplacement_ShouldWork(
            string sourceText, string searchText, string replacementText, string expected)
        {
            // Act
            var result = _service.ReplaceTextWithCapitalizationPreservation(sourceText, searchText, replacementText);

            // Assert
            result.Should().Be(expected);
        }

        [Theory]
        [InlineData("HELLO WORLD", "world", "universe", "HELLO UNIVERSE")]
        [InlineData("hello world", "WORLD", "universe", "hello universe")]
        [InlineData("Hello World", "world", "universe", "Hello Universe")]
        public void ReplaceTextWithCapitalizationPreservation_PreservesCapitalization_ShouldWork(
            string sourceText, string searchText, string replacementText, string expected)
        {
            // Act
            var result = _service.ReplaceTextWithCapitalizationPreservation(sourceText, searchText, replacementText);

            // Assert
            result.Should().Be(expected);
        }

        [Theory]
        [InlineData("Hello world   ", "world", "universe", "Hello universe   ")]
        [InlineData("Text with spaces  ", "spaces", "content", "Text with content  ")]
        public void ReplaceTextWithCapitalizationPreservation_PreservesTrailingWhitespace_ShouldWork(
            string sourceText, string searchText, string replacementText, string expected)
        {
            // Act
            var result = _service.ReplaceTextWithCapitalizationPreservation(sourceText, searchText, replacementText);

            // Assert
            result.Should().Be(expected);
        }

        [Theory]
        [InlineData("", "test", "replacement", "")]
        [InlineData("Hello world", "", "replacement", "Hello world")]
        [InlineData("Hello world", "test", "", "Hello world")]
        public void ReplaceTextWithCapitalizationPreservation_WithInvalidInput_ShouldReturnOriginal(
            string sourceText, string searchText, string replacementText, string expected)
        {
            // Act
            var result = _service.ReplaceTextWithCapitalizationPreservation(sourceText, searchText, replacementText);

            // Assert
            result.Should().Be(expected);
        }

        [Theory]
        [InlineData("The quick brown fox", "quick", "fast", "The fast brown fox")]
        [InlineData("QUICK BROWN FOX", "quick", "fast", "FAST BROWN FOX")]
        [InlineData("quick brown fox", "QUICK", "fast", "fast brown fox")]
        public void ReplaceTextWithCapitalizationPreservation_CaseInsensitiveMatching_ShouldWork(
            string sourceText, string searchText, string replacementText, string expected)
        {
            // Act
            var result = _service.ReplaceTextWithCapitalizationPreservation(sourceText, searchText, replacementText);

            // Assert
            result.Should().Be(expected);
        }

        [Fact]
        public void ReplaceTextWithCapitalizationPreservation_NoMatch_ShouldReturnOriginal()
        {
            // Arrange
            var sourceText = "Hello world";
            var searchText = "nonexistent";
            var replacementText = "replacement";

            // Act
            var result = _service.ReplaceTextWithCapitalizationPreservation(sourceText, searchText, replacementText);

            // Assert
            result.Should().Be(sourceText);
        }

        [Theory]
        [InlineData("Document Title", "document", "file", "File Title")]
        [InlineData("Document TITLE", "title", "header", "Document HEADER")]
        [InlineData("document title", "Document", "file", "file title")]
        public void ReplaceTextWithCapitalizationPreservation_MixedCases_ShouldPreservePattern(
            string sourceText, string searchText, string replacementText, string expected)
        {
            // Act
            var result = _service.ReplaceTextWithCapitalizationPreservation(sourceText, searchText, replacementText);

            // Assert
            result.Should().Be(expected);
        }
    }
}