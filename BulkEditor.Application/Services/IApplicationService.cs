using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Application.Services
{
    /// <summary>
    /// Main application service interface for orchestrating document processing operations
    /// </summary>
    public interface IApplicationService
    {
        /// <summary>
        /// Processes a single document with progress reporting
        /// </summary>
        Task<Document> ProcessSingleDocumentAsync(string filePath, IProgress<string>? progress = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// Processes multiple documents in batch with progress reporting
        /// </summary>
        Task<IEnumerable<Document>> ProcessDocumentsBatchAsync(IEnumerable<string> filePaths, IProgress<BulkEditor.Core.Interfaces.BatchProcessingProgress>? progress = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// Validates selected files before processing
        /// </summary>
        Task<ValidationResult> ValidateFilesAsync(IEnumerable<string> filePaths);

        /// <summary>
        /// Gets processing statistics for completed operations
        /// </summary>
        ProcessingStatistics GetProcessingStatistics(IEnumerable<Document> documents);

        /// <summary>
        /// Exports processing results to various formats
        /// </summary>
        Task<bool> ExportResultsAsync(IEnumerable<Document> documents, string outputPath, ExportFormat format, CancellationToken cancellationToken = default);
    }

    /// <summary>
    /// File validation result
    /// </summary>
    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public List<string> ValidFiles { get; set; } = new();
        public List<string> InvalidFiles { get; set; } = new();
        public List<string> ErrorMessages { get; set; } = new();
    }

    /// <summary>
    /// Processing statistics summary
    /// </summary>
    public class ProcessingStatistics
    {
        public int TotalDocuments { get; set; }
        public int SuccessfulDocuments { get; set; }
        public int FailedDocuments { get; set; }
        public int TotalHyperlinks { get; set; }
        public int UpdatedHyperlinks { get; set; }
        public int ExpiredHyperlinks { get; set; }
        public int InvalidHyperlinks { get; set; }
        public TimeSpan TotalProcessingTime { get; set; }
        public double SuccessRate => TotalDocuments > 0 ? (double)SuccessfulDocuments / TotalDocuments * 100 : 0;
    }

    /// <summary>
    /// Export format options
    /// </summary>
    public enum ExportFormat
    {
        Excel,
        Csv,
        Json,
        Xml
    }
}