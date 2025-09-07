using System;
using System.Collections.Generic;

namespace BulkEditor.Core.Models
{
    /// <summary>
    /// Represents a document processing session with metadata and statistics
    /// </summary>
    public class ProcessingSession
    {
        public Guid SessionId { get; set; } = Guid.NewGuid();
        public DateTime StartTime { get; set; } = DateTime.UtcNow;
        public DateTime? EndTime { get; set; }
        public int TotalDocuments { get; set; }
        public int ProcessedDocuments { get; set; }
        public int SuccessfulDocuments { get; set; }
        public int FailedDocuments { get; set; }
        public TimeSpan? TotalProcessingTime { get; set; }
        public string Status { get; set; } = "In Progress";
        public string? ErrorMessage { get; set; }
        public Dictionary<string, object> Metadata { get; set; } = new();
    }
}