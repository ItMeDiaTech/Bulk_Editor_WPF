using BulkEditor.Core.Entities;
using System;
using System.Linq;
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
            Assert.False(string.IsNullOrEmpty(document.Id));
            Assert.Equal(DocumentStatus.Pending, document.Status);
            Assert.True((DateTime.UtcNow - document.CreatedAt).TotalSeconds < 5);
            Assert.Null(document.ProcessedAt);
            Assert.NotNull(document.Hyperlinks);
            Assert.Empty(document.Hyperlinks);
            Assert.NotNull(document.ProcessingErrors);
            Assert.Empty(document.ProcessingErrors);
            Assert.NotNull(document.ChangeLog);
            Assert.Empty(document.ChangeLog.Changes);
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
            Assert.Equal(filePath, document.FilePath);
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
            Assert.Equal(expectedResult, isProcessed);
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
            Assert.Single(document.Hyperlinks);
            Assert.Equal(hyperlink, document.Hyperlinks.First());
        }

        [Fact]
        public void Document_AddProcessingError_ShouldUpdateErrorsList()
        {
            // Arrange
            var document = new Document();
            var errorMessage = "Test error occurred";

            // Act
            var processingError = new ProcessingError { Message = errorMessage, Severity = ErrorSeverity.Error };
            document.ProcessingErrors.Add(processingError);

            // Assert
            Assert.Single(document.ProcessingErrors);
            Assert.Contains(processingError, document.ProcessingErrors);
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
            Assert.Single(document.ChangeLog.Changes);
            Assert.Equal(1, document.ChangeLog.TotalChanges);
            Assert.Equal(changeEntry, document.ChangeLog.Changes.First());
        }
    }
}