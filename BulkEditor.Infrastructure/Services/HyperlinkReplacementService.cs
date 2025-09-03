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

        public HyperlinkReplacementService(IHttpService httpService, ILoggingService logger)
        {
            _httpService = httpService ?? throw new ArgumentNullException(nameof(httpService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));

            // Regex for extracting 6-digit content IDs from Content ID field
            _contentIdRegex = new Regex(@"[0-9]{6}", RegexOptions.Compiled);
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
                if (string.IsNullOrWhiteSpace(contentId))
                    return string.Empty;

                // Extract 6-digit Content ID if it's in a different format
                var match = _contentIdRegex.Match(contentId);
                var cleanContentId = match.Success ? match.Value : contentId;

                // Simulate title lookup - in a real implementation, this would query a CMS or API
                // For now, we'll use a simple mapping or generate a title
                var title = await SimulateTitleLookupAsync(cleanContentId, cancellationToken);

                _logger.LogDebug("Looked up title '{Title}' for Content ID: {ContentId}", title, cleanContentId);
                return title;
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
                if (string.IsNullOrWhiteSpace(contentId))
                    throw new ArgumentException("Content ID cannot be null or empty", nameof(contentId));

                // Extract 6-digit Content ID if needed
                var match = _contentIdRegex.Match(contentId);
                var cleanContentId = match.Success ? match.Value : contentId;

                // Build URL using the same pattern as existing hyperlink processing
                var url = $"https://example.com/content/{cleanContentId}";

                _logger.LogDebug("Built URL '{Url}' from Content ID: {ContentId}", url, cleanContentId);
                return url;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error building URL from Content ID: {ContentId}", contentId);
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
                // Look up the title for the Content ID
                var newTitle = await LookupTitleByContentIdAsync(rule.ContentId, cancellationToken);
                if (string.IsNullOrEmpty(newTitle))
                {
                    result.ErrorMessage = $"Could not lookup title for Content ID: {rule.ContentId}";
                    return result;
                }

                // Extract 6-digit Content ID
                var match = _contentIdRegex.Match(rule.ContentId);
                var cleanContentId = match.Success ? match.Value : rule.ContentId;

                // Build new display text: "Title (Content_ID)"
                var newDisplayText = $"{newTitle} ({cleanContentId})";

                // Build new URL
                var newUrl = BuildUrlFromContentId(cleanContentId);

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
                    Details = $"Content ID: {cleanContentId}, URL: {newUrl}"
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

            // Remove patterns like " (123456)" or "(123456)" from the end of text
            var pattern = @"\s*\([0-9]{6}\)\s*$";
            var result = Regex.Replace(text, pattern, "", RegexOptions.IgnoreCase).Trim();

            return result;
        }

        private async Task<string> SimulateTitleLookupAsync(string contentId, CancellationToken cancellationToken)
        {
            try
            {
                // Simulate API call delay
                await Task.Delay(50, cancellationToken);

                // In a real implementation, this would query a CMS/API
                // For now, return a formatted title based on Content ID
                return $"Document Title {contentId}";
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Error in title lookup simulation for Content ID: {ContentId}. Error: {Error}", contentId, ex.Message);
                return $"Document {contentId}";
            }
        }
    }
}