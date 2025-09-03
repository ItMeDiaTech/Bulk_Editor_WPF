using System;
using System.Collections.Generic;

namespace BulkEditor.Core.Entities
{
    /// <summary>
    /// Represents the change log for a document processing operation
    /// </summary>
    public class ChangeLog
    {
        public List<ChangeEntry> Changes { get; set; } = new();
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public string Summary { get; set; } = string.Empty;
        public int TotalChanges => Changes.Count;
        public bool HasErrors => Changes.Exists(c => c.Type == ChangeType.Error);
    }

    /// <summary>
    /// Represents a single change entry in the change log
    /// </summary>
    public class ChangeEntry
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
        public ChangeType Type { get; set; }
        public string Description { get; set; } = string.Empty;
        public string OldValue { get; set; } = string.Empty;
        public string NewValue { get; set; } = string.Empty;
        public string ElementId { get; set; } = string.Empty; // Reference to hyperlink or other element
        public string Details { get; set; } = string.Empty;
    }

    public enum ChangeType
    {
        Information,
        HyperlinkUpdated,
        HyperlinkRemoved,
        ContentIdAdded,
        TitleChanged,
        TitleReplaced,
        PossibleTitleChange,
        TextOptimized,
        TextReplaced,
        Error,
        Warning
    }
}