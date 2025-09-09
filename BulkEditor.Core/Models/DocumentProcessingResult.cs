using System;
using System.Collections.Generic;

namespace BulkEditor.Core.Models
{
    /// <summary>
    /// Represents the result of processing a single document
    /// </summary>
    public class DocumentProcessingResult
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public Guid SessionId { get; set; }
        public string DocumentPath { get; set; } = string.Empty;
        public string DocumentName { get; set; } = string.Empty;
        public long FileSizeBytes { get; set; }
        public DateTime ProcessingStartTime { get; set; }
        public DateTime ProcessingEndTime { get; set; }
        public TimeSpan ProcessingDuration { get; set; }
        public bool IsSuccessful { get; set; }
        public string? ErrorMessage { get; set; }
        public int HyperlinksProcessed { get; set; }
        public int HyperlinksUpdated { get; set; }
        public int TextReplacements { get; set; }
        public Dictionary<string, object> Metadata { get; set; } = new();
    }
}