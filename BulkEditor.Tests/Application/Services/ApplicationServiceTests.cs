using BulkEditor.Application.Services;
using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using FluentAssertions;
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

            _mockFileService.Setup(x => x.FileExists(filePath)).Returns(true);
            _mockFileService.Setup(x => x.IsValidWordDocument(filePath)).Returns(true);
            _mockFileService.Setup(x => x.GetFileInfo(filePath)).Returns(new System.IO.FileInfo(filePath));
            _mockDocumentProcessor.Setup(x => x.ProcessDocumentAsync(filePath, It.IsAny<IProgress<string>>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(expectedDocument);

            // Act
            var result = await _applicationService.ProcessSingleDocumentAsync(filePath);

            // Assert
            result.Should().NotBeNull();
            result.FilePath.Should().Be(filePath);
            result.Status.Should().Be(DocumentStatus.Completed);
        }

        [Fact]
        public async Task ValidateFilesAsync_ValidFiles_ShouldReturnValidResult()
        {
            // Arrange
            var filePaths = new[] { "test1.docx", "test2.docx" };

            _mockFileService.Setup(x => x.FileExists(It.IsAny<string>())).Returns(true);
            _mockFileService.Setup(x => x.IsValidWordDocument(It.IsAny<string>())).Returns(true);
            _mockFileService.Setup(x => x.GetFileInfo(It.IsAny<string>())).Returns(new System.IO.FileInfo("test.docx"));

            // Act
            var result = await _applicationService.ValidateFilesAsync(filePaths);

            // Assert
            result.Should().NotBeNull();
            result.ValidFiles.Should().HaveCount(2);
            result.InvalidFiles.Should().BeEmpty();
            result.ErrorMessages.Should().BeEmpty();
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
            result.Should().NotBeNull();
            result.ValidFiles.Should().BeEmpty();
            result.InvalidFiles.Should().HaveCount(2);
            result.ErrorMessages.Should().HaveCount(2);
            result.IsValid.Should().BeFalse();
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
            stats.Should().NotBeNull();
            stats.TotalDocuments.Should().Be(2);
            stats.SuccessfulDocuments.Should().Be(1);
            stats.FailedDocuments.Should().Be(1);
            stats.TotalHyperlinks.Should().Be(3);
            stats.UpdatedHyperlinks.Should().Be(1);
            stats.SuccessRate.Should().Be(50.0);
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

            _mockFileService.Setup(x => x.FileExists(It.IsAny<string>())).Returns(true);
            _mockFileService.Setup(x => x.IsValidWordDocument(It.IsAny<string>())).Returns(true);
            _mockFileService.Setup(x => x.GetFileInfo(It.IsAny<string>())).Returns(new System.IO.FileInfo("test.docx"));
            _mockDocumentProcessor.Setup(x => x.ProcessDocumentsBatchAsync(It.IsAny<IEnumerable<string>>(), It.IsAny<IProgress<BatchProcessingProgress>>(), It.IsAny<CancellationToken>()))
                .ReturnsAsync(expectedDocuments);

            // Act
            var results = await _applicationService.ProcessDocumentsBatchAsync(filePaths);

            // Assert
            results.Should().NotBeNull();
            results.Should().HaveCount(2);
            results.All(d => d.Status == DocumentStatus.Completed).Should().BeTrue();
        }

    }
}