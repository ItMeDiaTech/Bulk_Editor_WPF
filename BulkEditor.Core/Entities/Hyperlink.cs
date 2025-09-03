using System;

namespace BulkEditor.Core.Entities
{
    /// <summary>
    /// Represents a hyperlink found in a document
    /// </summary>
    public class Hyperlink
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string DisplayText { get; set; } = string.Empty;
        public string OriginalUrl { get; set; } = string.Empty;
        public string UpdatedUrl { get; set; } = string.Empty;
        public HyperlinkStatus Status { get; set; } = HyperlinkStatus.Pending;
        public string LookupId { get; set; } = string.Empty; // TSRC-xxx-xxxxxx or CMS-xxx-xxxxxx
        public string ContentId { get; set; } = string.Empty; // Used for title display (6-digit with leading zero padding)
        public string DocumentId { get; set; } = string.Empty; // Used for URL generation in docid parameter
        public DateTime? LastChecked { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
        public bool RequiresUpdate { get; set; }
        public HyperlinkAction ActionTaken { get; set; } = HyperlinkAction.None;
    }

    public enum HyperlinkStatus
    {
        Pending,
        Valid,
        Invalid,
        NotFound,
        Expired,
        Error
    }

    public enum HyperlinkAction
    {
        None,
        Updated,
        Removed,
        ContentIdAdded,
        UrlCorrected
    }
}