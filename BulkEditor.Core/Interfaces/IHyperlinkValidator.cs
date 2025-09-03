using BulkEditor.Core.Entities;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Core.Interfaces
{
    /// <summary>
    /// Interface for hyperlink validation operations
    /// </summary>
    public interface IHyperlinkValidator
    {
        /// <summary>
        /// Validates a single hyperlink asynchronously
        /// </summary>
        Task<HyperlinkValidationResult> ValidateHyperlinkAsync(Hyperlink hyperlink, CancellationToken cancellationToken = default);

        /// <summary>
        /// Validates multiple hyperlinks concurrently
        /// </summary>
        Task<IEnumerable<HyperlinkValidationResult>> ValidateHyperlinksAsync(IEnumerable<Hyperlink> hyperlinks, CancellationToken cancellationToken = default);

        /// <summary>
        /// Extracts lookup ID from hyperlink URL using regex pattern
        /// </summary>
        string ExtractLookupId(string url);

        /// <summary>
        /// Checks if a URL is expired based on content or response
        /// </summary>
        Task<bool> IsUrlExpiredAsync(string url, CancellationToken cancellationToken = default);

        /// <summary>
        /// Generates Content ID for a hyperlink based on lookup ID
        /// </summary>
        Task<string> GenerateContentIdAsync(string lookupId, CancellationToken cancellationToken = default);
    }

    /// <summary>
    /// Result of hyperlink validation
    /// </summary>
    public class HyperlinkValidationResult
    {
        public string HyperlinkId { get; set; } = string.Empty;
        public HyperlinkStatus Status { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
        public string LookupId { get; set; } = string.Empty;
        public string ContentId { get; set; } = string.Empty;
        public bool IsExpired { get; set; }
        public bool RequiresUpdate { get; set; }
        public string SuggestedUrl { get; set; } = string.Empty;
        public Entities.TitleComparisonResult? TitleComparison { get; set; }
    }
}