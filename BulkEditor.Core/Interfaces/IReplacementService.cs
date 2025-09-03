using BulkEditor.Core.Entities;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Core.Interfaces
{
    /// <summary>
    /// Main interface for replacement operations
    /// </summary>
    public interface IReplacementService
    {
        /// <summary>
        /// Processes all replacements for a document
        /// </summary>
        Task<Entities.Document> ProcessReplacementsAsync(Entities.Document document, CancellationToken cancellationToken = default);

        /// <summary>
        /// Validates replacement rules for correctness
        /// </summary>
        Task<ReplacementValidationResult> ValidateReplacementRulesAsync(IEnumerable<object> rules, CancellationToken cancellationToken = default);
    }

    /// <summary>
    /// Interface for hyperlink replacement operations
    /// </summary>
    public interface IHyperlinkReplacementService
    {
        /// <summary>
        /// Processes hyperlink replacements in a document
        /// </summary>
        Task<Entities.Document> ProcessHyperlinkReplacementsAsync(Entities.Document document, IEnumerable<Configuration.HyperlinkReplacementRule> rules, CancellationToken cancellationToken = default);

        /// <summary>
        /// Looks up document title by Content ID
        /// </summary>
        Task<string> LookupTitleByContentIdAsync(string contentId, CancellationToken cancellationToken = default);

        /// <summary>
        /// Builds URL from Content ID using existing URL generation logic
        /// </summary>
        string BuildUrlFromContentId(string contentId);
    }

    /// <summary>
    /// Interface for text replacement operations
    /// </summary>
    public interface ITextReplacementService
    {
        /// <summary>
        /// Processes text replacements in a document
        /// </summary>
        Task<Entities.Document> ProcessTextReplacementsAsync(Entities.Document document, IEnumerable<Configuration.TextReplacementRule> rules, CancellationToken cancellationToken = default);

        /// <summary>
        /// Performs case-insensitive text matching with capitalization preservation
        /// </summary>
        string ReplaceTextWithCapitalizationPreservation(string sourceText, string searchText, string replacementText);
    }

    /// <summary>
    /// Result of replacement rule validation
    /// </summary>
    public class ReplacementValidationResult
    {
        public bool IsValid { get; set; }
        public List<string> ValidationErrors { get; set; } = new();
        public int ValidRulesCount { get; set; }
        public int InvalidRulesCount { get; set; }
    }

    /// <summary>
    /// Result of hyperlink replacement operation
    /// </summary>
    public class HyperlinkReplacementResult
    {
        public string HyperlinkId { get; set; } = string.Empty;
        public bool WasReplaced { get; set; }
        public string OriginalTitle { get; set; } = string.Empty;
        public string NewTitle { get; set; } = string.Empty;
        public string NewUrl { get; set; } = string.Empty;
        public string ContentId { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
    }

    /// <summary>
    /// Result of text replacement operation
    /// </summary>
    public class TextReplacementResult
    {
        public string ElementId { get; set; } = string.Empty;
        public bool WasReplaced { get; set; }
        public string OriginalText { get; set; } = string.Empty;
        public string ReplacedText { get; set; } = string.Empty;
        public int ReplacementCount { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
    }
}