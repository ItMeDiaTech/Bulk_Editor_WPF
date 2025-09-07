using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using BulkEditor.Core.Models;

namespace BulkEditor.Core.Services
{
    /// <summary>
    /// Database service for persistent data storage using SQLite
    /// </summary>
    public interface IDatabaseService
    {
        /// <summary>
        /// Initializes the database and creates tables if they don't exist
        /// </summary>
        Task InitializeAsync(CancellationToken cancellationToken = default);

        /// <summary>
        /// Stores settings with transactional support
        /// </summary>
        Task SaveSettingsAsync(string key, string value, string category = "General", CancellationToken cancellationToken = default);

        /// <summary>
        /// Retrieves settings value
        /// </summary>
        Task<string?> GetSettingsAsync(string key, string category = "General", CancellationToken cancellationToken = default);

        /// <summary>
        /// Gets all settings in a category
        /// </summary>
        Task<Dictionary<string, string>> GetAllSettingsAsync(string category = "General", CancellationToken cancellationToken = default);

        /// <summary>
        /// Stores document processing session information
        /// </summary>
        Task SaveProcessingSessionAsync(ProcessingSession session, CancellationToken cancellationToken = default);

        /// <summary>
        /// Gets processing session by ID
        /// </summary>
        Task<ProcessingSession?> GetProcessingSessionAsync(Guid sessionId, CancellationToken cancellationToken = default);

        /// <summary>
        /// Gets recent processing sessions
        /// </summary>
        Task<IEnumerable<ProcessingSession>> GetRecentProcessingSessionsAsync(int limit = 50, CancellationToken cancellationToken = default);

        /// <summary>
        /// Stores document processing result
        /// </summary>
        Task SaveDocumentProcessingResultAsync(DocumentProcessingResult result, CancellationToken cancellationToken = default);

        /// <summary>
        /// Gets document processing history
        /// </summary>
        Task<IEnumerable<DocumentProcessingResult>> GetDocumentProcessingHistoryAsync(string? documentPath = null, int limit = 100, CancellationToken cancellationToken = default);

        /// <summary>
        /// Stores performance metrics
        /// </summary>
        Task SavePerformanceMetricAsync(PerformanceMetric metric, CancellationToken cancellationToken = default);

        /// <summary>
        /// Gets performance metrics for analysis
        /// </summary>
        Task<IEnumerable<PerformanceMetric>> GetPerformanceMetricsAsync(string? operationName = null, DateTime? fromDate = null, int limit = 1000, CancellationToken cancellationToken = default);

        /// <summary>
        /// Stores persistent cache entries
        /// </summary>
        Task SaveCacheEntryAsync(string key, string value, DateTime? expiryDate = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// Gets cache entry value
        /// </summary>
        Task<string?> GetCacheEntryAsync(string key, CancellationToken cancellationToken = default);

        /// <summary>
        /// Removes expired cache entries
        /// </summary>
        Task CleanupExpiredCacheAsync(CancellationToken cancellationToken = default);

        /// <summary>
        /// Gets database statistics
        /// </summary>
        Task<DatabaseStatistics> GetStatisticsAsync(CancellationToken cancellationToken = default);

        /// <summary>
        /// Executes database maintenance operations
        /// </summary>
        Task PerformMaintenanceAsync(CancellationToken cancellationToken = default);
    }
}