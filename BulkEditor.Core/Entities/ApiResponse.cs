using System.Collections.Generic;

namespace BulkEditor.Core.Entities
{
    /// <summary>
    /// Represents API response for hyperlink validation and title lookup
    /// </summary>
    public class ApiResponse
    {
        public string Version { get; set; } = string.Empty;
        public string Changes { get; set; } = string.Empty;
        public List<DocumentRecord> Results { get; set; } = new();
    }

    /// <summary>
    /// Represents a document record from API response
    /// </summary>
    public class DocumentRecord
    {
        public string Document_ID { get; set; } = string.Empty;
        public string Content_ID { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string Lookup_ID { get; set; } = string.Empty;
    }

    /// <summary>
    /// Result of title comparison operation
    /// </summary>
    public class TitleComparisonResult
    {
        public bool TitlesDiffer { get; set; }
        public string CurrentTitle { get; set; } = string.Empty;
        public string ApiTitle { get; set; } = string.Empty;
        public string ContentId { get; set; } = string.Empty;
        public bool WasReplaced { get; set; }
        public string ActionTaken { get; set; } = string.Empty;
    }
}