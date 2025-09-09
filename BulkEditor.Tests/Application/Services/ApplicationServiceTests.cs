using BulkEditor.Application.Services;
using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using Moq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace BulkEditor.Tests.Application.Services
{
    /// <summary>
    /// Unit tests for ApplicationService
    /// </summary>
    public class ApplicationServiceTests
    {
        private readonly Mock<IDocumentProcessor> _mockDocumentProcessor;
        private readonly Mock<IFileService> _mockFileService;
        private readonly Mock<ILoggingService> _mockLogger;
        private readonly ApplicationService _applicationService;

        public ApplicationServiceTests()
        {
            _mockDocumentProcessor = new Mock<IDocumentProcessor>();
            _mockFileService = new Mock<IFileService>();
            _mockLogger = new Mock<ILoggingService>();

            _applicationService = new ApplicationService(
                _mockDocumentProcessor.Object,
                _mockFileService.Object,
                _mockLogger.Object);
        }

        [Fact]
        public async Task ProcessSingleDocumentAsync_ValidFile_ShouldReturnCompletedDocument()
        {
            // Arrange
            var filePath = "test.docx";
            var expectedDocument = new Document
            {
                FilePath = filePath,
                FileName = "test.docx",
                Status = DocumentStatus.Completed
            };

            // Create a temporary file for testing
            var tempFile = Path.GetTempFileName();
            File.WriteAllText(tempFile, "test content");
            var fileInfo = new FileInfo(tempFile);
            fileInfo.IsReadOnly = false;

            try
            {
                _mockFileService.Setup(x => x.FileExists(filePath)).Returns(true);
                _mockFileService.Setup(x => x.IsValidWordDocument(filePath)).Returns(true);
                _mockFileService.Setup(x => x.GetFileInfo(filePath)).Returns(fileInfo);
                _mockDocumentProcessor.Setup(x => x.ProcessDocumentAsync(filePath, It.IsAny<IProgress<string>>(), It.IsAny<CancellationToken>()))
                    .ReturnsAsync(expectedDocument);

                // Act
                var result = await _applicationService.ProcessSingleDocumentAsync(filePath);

                // Assert
                Assert.NotNull(result);
                Assert.Equal(filePath, result.FilePath);
                Assert.Equal(DocumentStatus.Completed, result.Status);
            }
            finally
            {
                // Cleanup
                if (File.Exists(tempFile))
                    File.Delete(tempFile);
            }
        }

        [Fact]
        public async Task ValidateFilesAsync_ValidFiles_ShouldReturnValidResult()
        {
            // Arrange
            var filePaths = new[] { "test1.docx", "test2.docx" };

            // Create a temporary file for testing
            var tempFile = Path.GetTempFileName();
            File.WriteAllText(tempFile, "test content");
            var fileInfo = new FileInfo(tempFile);
            fileInfo.IsReadOnly = false;

            try
            {
                _mockFileService.Setup(x => x.FileExists(It.IsAny<string>())).Returns(true);
                _mockFileService.Setup(x => x.IsValidWordDocument(It.IsAny<string>())).Returns(true);
                _mockFileService.Setup(x => x.GetFileInfo(It.IsAny<string>())).Returns(fileInfo);

                // Act
                var result = await _applicationService.ValidateFilesAsync(filePaths);

                // Assert
                Assert.NotNull(result);
                Assert.Equal(2, result.ValidFiles.Count);
                Assert.Empty(result.InvalidFiles);
                Assert.Empty(result.ErrorMessages);
            }
            finally
            {
                // Cleanup
                if (File.Exists(tempFile))
                    File.Delete(tempFile);
            }
        }

        [Fact]
        public async Task ValidateFilesAsync_InvalidFiles_ShouldReturnInvalidResults()
        {
            // Arrange
            var filePaths = new[] { "invalid.txt", "missing.docx" };

            _mockFileService.Setup(x => x.FileExists("invalid.txt")).Returns(true);
            _mockFileService.Setup(x => x.IsValidWordDocument("invalid.txt")).Returns(false);
            _mockFileService.Setup(x => x.FileExists("missing.docx")).Returns(false);

            // Act
            var result = await _applicationService.ValidateFilesAsync(filePaths);

            // Assert
            Assert.NotNull(result);
            Assert.Empty(result.ValidFiles);
            Assert.Equal(2, result.InvalidFiles.Count);
            Assert.Equal(2, result.ErrorMessages.Count);
            Assert.False(result.IsValid);
        }

        [Fact]
        public void GetProcessingStatistics_CompletedDocuments_ShouldCalculateCorrectStats()
        {
            // Arrange
            var documents = new List<Document>
            {
                new Document
                {
                    Status = DocumentStatus.Completed,
                    Hyperlinks = new List<Hyperlink>
                    {
                        new Hyperlink { ActionTaken = HyperlinkAction.Updated },
                        new Hyperlink { Status = HyperlinkStatus.Valid }
                    }
                },
                new Document
                {
                    Status = DocumentStatus.Failed,
                    Hyperlinks = new List<Hyperlink>
                    {
                        new Hyperlink { Status = HyperlinkStatus.Invalid }
                    }
                }
            };

            // Act
            var stats = _applicationService.GetProcessingStatistics(documents);

            // Assert
            Assert.NotNull(stats);
            Assert.Equal(2, stats.TotalDocuments);
            Assert.Equal(1, stats.SuccessfulDocuments);
            Assert.Equal(1, stats.FailedDocuments);
            Assert.Equal(3, stats.TotalHyperlinks);
            Assert.Equal(1, stats.UpdatedHyperlinks);
            Assert.Equal(50.0, stats.SuccessRate);
        }

        [Fact]
        public async Task ProcessDocumentsBatchAsync_MultipleFiles_ShouldProcessAll()
        {
            // Arrange
            var filePaths = new[] { "doc1.docx", "doc2.docx" };
            var expectedDocuments = filePaths.Select(path => new Document
            {
                FilePath = path,
                FileName = Path.GetFileName(path),
                Status = DocumentStatus.Completed
            }).ToList();

            // Create a temporary file for testing
            var tempFile = Path.GetTempFileName();
            File.WriteAllText(tempFile, "test content");
            var fileInfo = new FileInfo(tempFile);
            fileInfo.IsReadOnly = false;

            try
            {
                _mockFileService.Setup(x => x.FileExists(It.IsAny<string>())).Returns(true);
                _mockFileService.Setup(x => x.IsValidWordDocument(It.IsAny<string>())).Returns(true);
                _mockFileService.Setup(x => x.GetFileInfo(It.IsAny<string>())).Returns(fileInfo);
                _mockDocumentProcessor.Setup(x => x.ProcessDocumentsBatchAsync(It.IsAny<IEnumerable<string>>(), It.IsAny<IProgress<BatchProcessingProgress>>(), It.IsAny<CancellationToken>()))
                    .ReturnsAsync(expectedDocuments);

                // Act
                var results = await _applicationService.ProcessDocumentsBatchAsync(filePaths);

                // Assert
                Assert.NotNull(results);
                Assert.Equal(2, results.Count());
                Assert.True(results.All(d => d.Status == DocumentStatus.Completed));
            }
            finally
            {
                // Cleanup
                if (File.Exists(tempFile))
                    File.Delete(tempFile);
            }
        }

    }
}