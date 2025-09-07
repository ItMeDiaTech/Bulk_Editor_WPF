using BulkEditor.Core.Entities;
using DocumentFormat.OpenXml.Packaging;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Core.Interfaces
{
    /// <summary>
    /// Interface for text optimization operations
    /// </summary>
    public interface ITextOptimizer
    {
        /// <summary>
        /// Optimizes text formatting in a document using an already opened WordprocessingDocument
        /// </summary>
        /// <param name="wordDocument">Already opened WordprocessingDocument</param>
        /// <param name="document">Document entity to track changes</param>
        /// <param name="cancellationToken">Cancellation token</param>
        /// <returns>Number of optimizations made</returns>
        Task<int> OptimizeDocumentTextInSessionAsync(WordprocessingDocument wordDocument, Document document, CancellationToken cancellationToken = default);

        /// <summary>
        /// Optimizes text formatting in a document (legacy method - opens document independently)
        /// </summary>
        [System.Obsolete("Use OptimizeDocumentTextInSessionAsync to prevent file corruption")]
        Task<Document> OptimizeDocumentTextAsync(Document document, CancellationToken cancellationToken = default);

        /// <summary>
        /// Removes extra spaces and standardizes whitespace
        /// </summary>
        Task<string> OptimizeWhitespaceAsync(string text, CancellationToken cancellationToken = default);

        /// <summary>
        /// Standardizes paragraph formatting
        /// </summary>
        Task<Document> StandardizeParagraphFormattingAsync(Document document, CancellationToken cancellationToken = default);

        /// <summary>
        /// Removes unnecessary formatting elements
        /// </summary>
        Task<Document> RemoveUnnecessaryFormattingAsync(Document document, CancellationToken cancellationToken = default);

        /// <summary>
        /// Optimizes document structure and hierarchy
        /// </summary>
        Task<Document> OptimizeDocumentStructureAsync(Document document, CancellationToken cancellationToken = default);
    }

    /// <summary>
    /// Text optimization settings
    /// </summary>
    public class TextOptimizationSettings
    {
        /// <summary>
        /// Remove extra spaces between words
        /// </summary>
        public bool RemoveExtraSpaces { get; set; } = true;

        /// <summary>
        /// Standardize line breaks and paragraph spacing
        /// </summary>
        public bool StandardizeLineBreaks { get; set; } = true;

        /// <summary>
        /// Remove empty paragraphs
        /// </summary>
        public bool RemoveEmptyParagraphs { get; set; } = true;

        /// <summary>
        /// Standardize heading formatting
        /// </summary>
        public bool StandardizeHeadings { get; set; } = true;

        /// <summary>
        /// Remove unnecessary text formatting
        /// </summary>
        public bool RemoveUnnecessaryFormatting { get; set; } = false;

        /// <summary>
        /// Optimize table formatting
        /// </summary>
        public bool OptimizeTableFormatting { get; set; } = true;

        /// <summary>
        /// Clean up list formatting
        /// </summary>
        public bool OptimizeListFormatting { get; set; } = true;

        /// <summary>
        /// Maximum number of consecutive line breaks allowed
        /// </summary>
        public int MaxConsecutiveLineBreaks { get; set; } = 2;
    }
}