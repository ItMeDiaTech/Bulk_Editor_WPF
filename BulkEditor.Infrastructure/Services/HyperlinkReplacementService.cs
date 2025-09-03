using BulkEditor.Core.Configuration;
using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using OpenXmlHyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;
using CoreDocument = BulkEditor.Core.Entities.Document;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Implementation of hyperlink replacement service
    /// </summary>
    public class HyperlinkReplacementService : IHyperlinkReplacementService
    {
        private readonly IHttpService _httpService;
        private readonly ILoggingService _logger;
        private readonly Regex _contentIdRegex;
        private readonly Regex _fiveDigitRegex;
        private readonly Regex _sixDigitRegex;

        public HyperlinkReplacementService(IHttpService httpService, ILoggingService logger)
        {
            _httpService = httpService ?? throw new ArgumentNullException(nameof(httpService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));

            // Regex for extracting 6-digit content IDs from Content ID field
            _contentIdRegex = new Regex(@"[0-9]{6}", RegexOptions.Compiled);
            // Regex for detecting 5-digit content IDs that need padding
            _fiveDigitRegex = new Regex(@"[0-9]{5}", RegexOptions.Compiled);
            // Regex for detecting 6-digit content IDs
            _sixDigitRegex = new Regex(@"[0-9]{6}", RegexOptions.Compiled);
        }

        public async Task<CoreDocument> ProcessHyperlinkReplacementsAsync(CoreDocument document, IEnumerable<HyperlinkReplacementRule> rules, CancellationToken cancellationToken = default)
        {
            try
            {
                var activeRules = rules.Where(r => r.IsEnabled && !string.IsNullOrWhiteSpace(r.TitleToMatch) && !string.IsNullOrWhiteSpace(r.ContentId)).ToList();

                if (!activeRules.Any())
                {
                    _logger.LogDebug("No active hyperlink replacement rules found for document: {FileName}", document.FileName);
                    return document;
                }

                _logger.LogInformation("Processing {Count} hyperlink replacement rules for document: {FileName}", activeRules.Count, document.FileName);

                using var wordDocument = WordprocessingDocument.Open(document.FilePath, true);
                var mainPart = wordDocument.MainDocumentPart;

                if (mainPart?.Document?.Body != null)
                {
                    var hyperlinks = mainPart.Document.Body.Descendants<OpenXmlHyperlink>().ToList();
                    var replacementsMade = 0;

                    foreach (var openXmlHyperlink in hyperlinks)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        var currentDisplayText = openXmlHyperlink.InnerText?.Trim();
                        if (string.IsNullOrEmpty(currentDisplayText))
                            continue;

                        // Remove any existing Content ID from display text for comparison
                        var cleanDisplayText = RemoveContentIdFromText(currentDisplayText).Trim().ToLowerInvariant();

                        foreach (var rule in activeRules)
                        {
                            var ruleTitleLower = rule.TitleToMatch.Trim().ToLowerInvariant();

                            if (cleanDisplayText.Equals(ruleTitleLower, StringComparison.OrdinalIgnoreCase))
                            {
                                var result = await ProcessHyperlinkReplacementAsync(mainPart, openXmlHyperlink, rule, document, cancellationToken);
                                if (result.WasReplaced)
                                {
                                    replacementsMade++;
                                    _logger.LogInformation("Replaced hyperlink: '{OriginalTitle}' -> '{NewTitle}' with Content ID: {ContentId}",
                                        result.OriginalTitle, result.NewTitle, result.ContentId);
                                }
                                break; // Only apply first matching rule
                            }
                        }
                    }

                    if (replacementsMade > 0)
                    {
                        // Save the document
                        mainPart.Document.Save();
                        _logger.LogInformation("Saved document with {Count} hyperlink replacements: {FileName}", replacementsMade, document.FileName);
                    }
                }

                return document;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing hyperlink replacements for document: {FileName}", document.FileName);
                throw;
            }
        }

        public async Task<string> LookupTitleByContentIdAsync(string contentId, CancellationToken cancellationToken = default)
        {
            try
            {
                var documentRecord = await LookupDocumentByContentIdAsync(contentId, cancellationToken);
                return documentRecord?.Title ?? $"Document {contentId}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error looking up title for Content ID: {ContentId}", contentId);
                return $"Document {contentId}"; // Fallback title
            }
        }

        public string BuildUrlFromContentId(string contentId)
        {
            try
            {
                // This is the legacy method that incorrectly used Content_ID for URL
                // For backward compatibility, we'll simulate the old behavior
                // but log a warning that this should use Document_ID instead
                _logger.LogWarning("BuildUrlFromContentId called - this method should use Document_ID instead of Content_ID for URL generation");

                if (string.IsNullOrWhiteSpace(contentId))
                    throw new ArgumentException("Content ID cannot be null or empty", nameof(contentId));

                // Extract 6-digit Content ID if needed
                var match = _contentIdRegex.Match(contentId);
                var cleanContentId = match.Success ? match.Value : contentId;

                // For backward compatibility, build URL using Content_ID (but this is incorrect)
                var url = $"https://thesource.cvshealth.com/nuxeo/thesource/#!/view?docid={cleanContentId}";

                _logger.LogDebug("Built URL '{Url}' from Content ID: {ContentId} (legacy method)", url, cleanContentId);
                return url;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error building URL from Content ID: {ContentId}", contentId);
                throw;
            }
        }

        public async Task<DocumentRecord> LookupDocumentByContentIdAsync(string contentId, CancellationToken cancellationToken = default)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(contentId))
                    return null;

                // Extract 6-digit Content ID if it's in a different format
                var match = _contentIdRegex.Match(contentId);
                var cleanContentId = match.Success ? match.Value : contentId;

                // Simulate document lookup - in a real implementation, this would query a CMS or API
                // For now, we'll use a simple mapping or generate a document record
                var documentRecord = await SimulateDocumentLookupAsync(cleanContentId, cancellationToken);

                _logger.LogDebug("Looked up document record for Content ID: {ContentId}", cleanContentId);
                return documentRecord;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error looking up document for Content ID: {ContentId}", contentId);
                return null;
            }
        }

        public string BuildUrlFromDocumentId(string documentId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(documentId))
                    throw new ArgumentException("Document ID cannot be null or empty", nameof(documentId));

                // Build URL using the correct CVS Health format with Document_ID
                var url = $"https://thesource.cvshealth.com/nuxeo/thesource/#!/view?docid={documentId}";

                _logger.LogDebug("Built URL '{Url}' from Document ID: {DocumentId}", url, documentId);
                return url;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error building URL from Document ID: {DocumentId}", documentId);
                throw;
            }
        }

        private async Task<HyperlinkReplacementResult> ProcessHyperlinkReplacementAsync(
            MainDocumentPart mainPart,
            OpenXmlHyperlink openXmlHyperlink,
            HyperlinkReplacementRule rule,
            Document document,
            CancellationToken cancellationToken)
        {
            var result = new HyperlinkReplacementResult
            {
                HyperlinkId = openXmlHyperlink.Id?.Value ?? Guid.NewGuid().ToString(),
                OriginalTitle = openXmlHyperlink.InnerText
            };

            try
            {
                // Look up the document record for the Content ID
                var documentRecord = await LookupDocumentByContentIdAsync(rule.ContentId, cancellationToken);
                if (documentRecord == null || string.IsNullOrEmpty(documentRecord.Title))
                {
                    result.ErrorMessage = $"Could not lookup document for Content ID: {rule.ContentId}";
                    return result;
                }

                // Extract and format Content ID - pad 5-digit IDs with leading zero
                var cleanContentId = FormatContentId(rule.ContentId);

                // Build new display text: "Title (Content_ID)" - using Content_ID in title
                var newDisplayText = $"{documentRecord.Title} ({cleanContentId})";

                // Build new URL using Document_ID, not Content_ID
                var newUrl = BuildUrlFromDocumentId(documentRecord.Document_ID);

                // Update the hyperlink display text
                openXmlHyperlink.RemoveAllChildren();
                openXmlHyperlink.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(newDisplayText));

                // Update the URL if hyperlink has a relationship ID
                var relId = openXmlHyperlink.Id?.Value;
                if (!string.IsNullOrEmpty(relId))
                {
                    try
                    {
                        // Delete old relationship and create new one
                        mainPart.DeleteReferenceRelationship(relId);
                        mainPart.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                            new Uri(newUrl), relId);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning("Could not update hyperlink URL for relationship {RelId}: {Error}", relId, ex.Message);
                    }
                }

                // Update result
                result.WasReplaced = true;
                result.NewTitle = newDisplayText;
                result.NewUrl = newUrl;
                result.ContentId = cleanContentId;

                // Log the change in document
                document.ChangeLog.Changes.Add(new ChangeEntry
                {
                    Type = ChangeType.HyperlinkUpdated,
                    Description = "Hyperlink replaced using replacement rule",
                    OldValue = result.OriginalTitle,
                    NewValue = newDisplayText,
                    ElementId = result.HyperlinkId,
                    Details = $"Content ID: {cleanContentId}, Document ID: {documentRecord.Document_ID}, URL: {newUrl}"
                });

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing hyperlink replacement for rule: {RuleId}", rule.Id);
                result.ErrorMessage = ex.Message;
                return result;
            }
        }

        private string RemoveContentIdFromText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            // Remove patterns like " (123456)" or "(123456)" from the end of text (6-digit)
            var pattern6 = @"\s*\([0-9]{6}\)\s*$";
            var result = Regex.Replace(text, pattern6, "", RegexOptions.IgnoreCase).Trim();

            // Also remove 5-digit patterns like " (12345)" from the end of text
            var pattern5 = @"\s*\([0-9]{5}\)\s*$";
            result = Regex.Replace(result, pattern5, "", RegexOptions.IgnoreCase).Trim();

            return result;
        }

        /// <summary>
        /// Formats Content ID to ensure 6-digit format with leading zero padding if needed
        /// </summary>
        /// <param name="contentId">Raw content ID</param>
        /// <returns>Formatted 6-digit content ID</returns>
        private string FormatContentId(string contentId)
        {
            if (string.IsNullOrWhiteSpace(contentId))
                return contentId;

            // First try to extract 6-digit ID
            var sixDigitMatch = _sixDigitRegex.Match(contentId);
            if (sixDigitMatch.Success)
            {
                return sixDigitMatch.Value;
            }

            // If no 6-digit found, try 5-digit and pad with leading zero
            var fiveDigitMatch = _fiveDigitRegex.Match(contentId);
            if (fiveDigitMatch.Success)
            {
                var paddedId = "0" + fiveDigitMatch.Value;
                _logger.LogDebug("Padded 5-digit Content ID '{OriginalId}' to 6-digit '{PaddedId}'", fiveDigitMatch.Value, paddedId);
                return paddedId;
            }

            // Return original if no numeric pattern found
            return contentId;
        }

        private async Task<DocumentRecord> SimulateDocumentLookupAsync(string contentId, CancellationToken cancellationToken)
        {
            try
            {
                // Simulate API call delay
                await Task.Delay(50, cancellationToken);

                // In a real implementation, this would query a CMS/API
                // For now, return a simulated document record
                return new DocumentRecord
                {
                    Document_ID = $"doc-{contentId}-{DateTime.Now.Ticks % 1000}",
                    Content_ID = contentId,
                    Title = $"Document Title {contentId}",
                    Status = "Released",
                    Lookup_ID = $"TSRC-{contentId}"
                };
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Error in document lookup simulation for Content ID: {ContentId}. Error: {Error}", contentId, ex.Message);
                return new DocumentRecord
                {
                    Document_ID = $"fallback-doc-{contentId}",
                    Content_ID = contentId,
                    Title = $"Document {contentId}",
                    Status = "Unknown"
                };
            }
        }
    }
}
