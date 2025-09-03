using System;
using System.Collections.Generic;

namespace BulkEditor.Core.Entities
{
    /// <summary>
    /// Contains metadata about a document
    /// </summary>
    public class DocumentMetadata
    {
        public long FileSizeBytes { get; set; }
        public DateTime LastModified { get; set; }
        public string Author { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string Keywords { get; set; } = string.Empty;
        public string Comments { get; set; } = string.Empty;
        public int WordCount { get; set; }
        public int PageCount { get; set; }
        public int HyperlinkCount { get; set; }
        public bool HasExpiredLinks { get; set; }
        public bool HasInvalidLinks { get; set; }
        public Dictionary<string, object> CustomProperties { get; set; } = new();
    }
}