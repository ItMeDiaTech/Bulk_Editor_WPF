using BulkEditor.Core.Entities;
using FluentAssertions;
using Xunit;

namespace BulkEditor.Tests.Core.Entities
{
    /// <summary>
    /// Unit tests for Document entity
    /// </summary>
    public class DocumentTests
    {
        [Fact]
        public void Document_DefaultConstruction_ShouldSetCorrectDefaults()
        {
            // Arrange & Act
            var document = new Document();

            // Assert
            document.Id.Should().NotBeEmpty();
            document.Status.Should().Be(DocumentStatus.Pending);
            document.CreatedAt.Should().BeCloseTo(DateTime.UtcNow, TimeSpan.FromSeconds(5));
            document.ProcessedAt.Should().BeNull();
            document.Hyperlinks.Should().NotBeNull().And.BeEmpty();
            document.ProcessingErrors.Should().NotBeNull().And.BeEmpty();
            document.ChangeLog.Should().NotBeNull();
            document.ChangeLog.Changes.Should().BeEmpty();
        }

        [Fact]
        public void Document_WithFilePath_ShouldExtractFileName()
        {
            // Arrange
            var filePath = @"C:\Documents\TestDocument.docx";

            // Act
            var document = new Document
            {
                FilePath = filePath
            };

            // Assert
            document.FilePath.Should().Be(filePath);
        }

        [Theory]
        [InlineData(DocumentStatus.Pending, false)]
        [InlineData(DocumentStatus.Processing, false)]
        [InlineData(DocumentStatus.Completed, true)]
        [InlineData(DocumentStatus.Failed, true)]
        public void Document_IsProcessed_ShouldReturnCorrectValue(DocumentStatus status, bool expectedResult)
        {
            // Arrange
            var document = new Document { Status = status };

            // Act
            var isProcessed = document.Status == DocumentStatus.Completed || document.Status == DocumentStatus.Failed;

            // Assert
            isProcessed.Should().Be(expectedResult);
        }

        [Fact]
        public void Document_AddHyperlink_ShouldUpdateCollection()
        {
            // Arrange
            var document = new Document();
            var hyperlink = new Hyperlink
            {
                OriginalUrl = "https://example.com",
                DisplayText = "Example Link"
            };

            // Act
            document.Hyperlinks.Add(hyperlink);

            // Assert
            document.Hyperlinks.Should().HaveCount(1);
            document.Hyperlinks.First().Should().Be(hyperlink);
        }

        [Fact]
        public void Document_AddProcessingError_ShouldUpdateErrorsList()
        {
            // Arrange
            var document = new Document();
            var errorMessage = "Test error occurred";

            // Act
            document.ProcessingErrors.Add(errorMessage);

            // Assert
            document.ProcessingErrors.Should().HaveCount(1);
            document.ProcessingErrors.Should().Contain(errorMessage);
        }

        [Fact]
        public void Document_ChangeLog_ShouldTrackChanges()
        {
            // Arrange
            var document = new Document();
            var changeEntry = new ChangeEntry
            {
                Type = ChangeType.HyperlinkUpdated,
                Description = "Updated hyperlink URL",
                OldValue = "old-url",
                NewValue = "new-url"
            };

            // Act
            document.ChangeLog.Changes.Add(changeEntry);

            // Assert
            document.ChangeLog.Changes.Should().HaveCount(1);
            document.ChangeLog.TotalChanges.Should().Be(1);
            document.ChangeLog.Changes.First().Should().Be(changeEntry);
        }
    }
}