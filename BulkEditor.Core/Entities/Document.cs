using System;
using System.Collections.Generic;

namespace BulkEditor.Core.Entities
{
    /// <summary>
    /// Represents a document being processed by the bulk editor
    /// </summary>
    public class Document
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string FilePath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public DocumentStatus Status { get; set; } = DocumentStatus.Pending;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime? ProcessedAt { get; set; }
        public List<Hyperlink> Hyperlinks { get; set; } = new();
        public List<ProcessingError> ProcessingErrors { get; set; } = new();
        public string BackupPath { get; set; } = string.Empty;
        public DocumentMetadata Metadata { get; set; } = new();
        public ChangeLog ChangeLog { get; set; } = new();
    }

    public enum DocumentStatus
    {
        Pending,
        Processing,
        Completed,
        Failed,
        Cancelled,
        Recovered
    }
}