namespace BulkEditor.Core.Entities
{
    /// <summary>
    /// Represents an error that occurred during document processing.
    /// </summary>
    public class ProcessingError
    {
        /// <summary>
        /// Gets or sets the ID of the rule that caused the error, if applicable.
        /// </summary>
        public string RuleId { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the error message.
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the severity of the error.
        /// </summary>
        public ErrorSeverity Severity { get; set; }
    }

    /// <summary>
    /// Defines the severity levels for processing errors.
    /// </summary>
    public enum ErrorSeverity
    {
        /// <summary>
        /// A non-critical error that does not prevent further processing.
        /// </summary>
        Warning,

        /// <summary>
        /// A critical error that may prevent further processing.
        /// </summary>
        Error
    }
}