using System;

namespace BulkEditor.Core.Models
{
    /// <summary>
    /// Represents database statistics and information
    /// </summary>
    public class DatabaseStatistics
    {
        public long DatabaseSizeBytes { get; set; }
        public int SettingsCount { get; set; }
        public int ProcessingSessionsCount { get; set; }
        public int DocumentResultsCount { get; set; }
        public int PerformanceMetricsCount { get; set; }
        public int CacheEntriesCount { get; set; }
        public DateTime LastMaintenanceDate { get; set; }
    }
}