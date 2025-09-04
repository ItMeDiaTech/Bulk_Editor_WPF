using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Core.Interfaces
{
    /// <summary>
    /// Interface for document processing operations
    /// </summary>
    public interface IDocumentProcessor
    {
        /// <summary>
        /// Processes a single document asynchronously
        /// </summary>
        Task<Entities.Document> ProcessDocumentAsync(string filePath, IProgress<string>? progress = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// Processes multiple documents in batch
        /// </summary>
        Task<IEnumerable<Entities.Document>> ProcessDocumentsBatchAsync(IEnumerable<string> filePaths, IProgress<BatchProcessingProgress>? progress = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// Validates hyperlinks in a document
        /// </summary>
        Task<IEnumerable<Entities.Hyperlink>> ValidateHyperlinksAsync(Entities.Document document, CancellationToken cancellationToken = default);

        /// <summary>
        /// Creates a backup of the document before processing
        /// </summary>
        Task<string> CreateBackupAsync(string filePath, CancellationToken cancellationToken = default);

        /// <summary>
        /// Restores a document from backup
        /// </summary>
        Task<bool> RestoreFromBackupAsync(string filePath, string backupPath, CancellationToken cancellationToken = default);
    }

    /// <summary>
    /// Represents progress information for batch processing
    /// </summary>
    public class BatchProcessingProgress
    {
        public int TotalDocuments { get; set; }
        public int ProcessedDocuments { get; set; }
        public int FailedDocuments { get; set; }
        public string CurrentDocument { get; set; } = string.Empty;
        public double PercentageComplete => TotalDocuments > 0 ? (double)ProcessedDocuments / TotalDocuments * 100 : 0;
    }
}