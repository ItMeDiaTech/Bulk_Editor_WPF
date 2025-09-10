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

        // NOTE: Backup creation is now handled by IBackupService for centralized backup management

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
        public int SuccessfulDocuments { get; set; }
        public string CurrentDocument { get; set; } = string.Empty;
        public string CurrentOperation { get; set; } = string.Empty;
        public DateTime StartTime { get; set; }
        public DateTime? EstimatedCompletionTime { get; set; }
        public TimeSpan ElapsedTime => DateTime.Now - StartTime;
        public TimeSpan? EstimatedTimeRemaining { get; set; }
        public int TotalHyperlinksFound { get; set; }
        public int TotalHyperlinksProcessed { get; set; }
        public int TotalHyperlinksUpdated { get; set; }
        public int TotalTextReplacements { get; set; }
        public long TotalBytesProcessed { get; set; }
        public double AverageProcessingTimePerDocument { get; set; }
        public double CurrentDocumentProgress { get; set; }
        public List<string> RecentErrors { get; set; } = new();
        
        // CRITICAL FIX: Track unique hyperlinks that have been changed
        // This ensures each hyperlink is counted only once, regardless of how many changes were made
        public HashSet<string> UniqueHyperlinksChanged { get; set; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        
        public double PercentageComplete => TotalDocuments > 0 ? (double)ProcessedDocuments / TotalDocuments * 100 : 0;
        public double OverallProgress => TotalDocuments > 0 ? ((double)ProcessedDocuments + (CurrentDocumentProgress / 100.0)) / TotalDocuments * 100 : 0;
        public string FormattedElapsedTime => ElapsedTime.ToString(@"hh\:mm\:ss");
        public string FormattedEstimatedTimeRemaining => EstimatedTimeRemaining?.ToString(@"hh\:mm\:ss") ?? "Calculating...";
    }
}