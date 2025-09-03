using BulkEditor.Core.Configuration;
using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using CoreDocument = BulkEditor.Core.Entities.Document;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Implementation of text replacement service with capitalization preservation
    /// </summary>
    public class TextReplacementService : ITextReplacementService
    {
        private readonly ILoggingService _logger;

        public TextReplacementService(ILoggingService logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public async Task<CoreDocument> ProcessTextReplacementsAsync(CoreDocument document, IEnumerable<TextReplacementRule> rules, CancellationToken cancellationToken = default)
        {
            try
            {
                var activeRules = rules.Where(r => r.IsEnabled && !string.IsNullOrWhiteSpace(r.SourceText) && !string.IsNullOrWhiteSpace(r.ReplacementText)).ToList();

                if (!activeRules.Any())
                {
                    _logger.LogDebug("No active text replacement rules found for document: {FileName}", document.FileName);
                    return document;
                }

                _logger.LogInformation("Processing {Count} text replacement rules for document: {FileName}", activeRules.Count, document.FileName);

                using var wordDocument = WordprocessingDocument.Open(document.FilePath, true);
                var mainPart = wordDocument.MainDocumentPart;

                if (mainPart?.Document?.Body != null)
                {
                    var totalReplacements = 0;

                    // Process all text elements in the document
                    var textElements = mainPart.Document.Body.Descendants<Text>().ToList();

                    foreach (var textElement in textElements)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        if (string.IsNullOrEmpty(textElement.Text))
                            continue;

                        var originalText = textElement.Text;
                        var modifiedText = originalText;
                        var replacementsMadeInElement = 0;

                        foreach (var rule in activeRules)
                        {
                            var replacementResult = ReplaceTextWithCapitalizationPreservation(
                                modifiedText, rule.SourceText, rule.ReplacementText);

                            if (replacementResult != modifiedText)
                            {
                                modifiedText = replacementResult;
                                replacementsMadeInElement++;
                            }
                        }

                        // Update text element if any replacements were made
                        if (replacementsMadeInElement > 0)
                        {
                            textElement.Text = modifiedText;
                            totalReplacements += replacementsMadeInElement;

                            // Log the change in document
                            document.ChangeLog.Changes.Add(new ChangeEntry
                            {
                                Type = ChangeType.TextReplaced,
                                Description = "Text replaced using replacement rules",
                                OldValue = originalText,
                                NewValue = modifiedText,
                                ElementId = Guid.NewGuid().ToString(),
                                Details = $"Applied {replacementsMadeInElement} replacement rule(s)"
                            });

                            _logger.LogDebug("Text replacement in element: '{OriginalText}' -> '{NewText}'", originalText, modifiedText);
                        }
                    }

                    if (totalReplacements > 0)
                    {
                        // Save the document
                        mainPart.Document.Save();
                        _logger.LogInformation("Saved document with {Count} text replacements: {FileName}", totalReplacements, document.FileName);
                    }
                }

                // Simulate async operation
                await Task.CompletedTask;
                return document;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing text replacements for document: {FileName}", document.FileName);
                throw;
            }
        }

        public string ReplaceTextWithCapitalizationPreservation(string sourceText, string searchText, string replacementText)
        {
            try
            {
                if (string.IsNullOrEmpty(sourceText) || string.IsNullOrWhiteSpace(searchText) || string.IsNullOrWhiteSpace(replacementText))
                    return sourceText;

                // Trim trailing whitespace from search text for comparison
                var trimmedSearchText = searchText.TrimEnd();
                var trimmedSourceText = sourceText.TrimEnd();

                // Create regex pattern for case-insensitive matching
                var escapedSearchText = Regex.Escape(trimmedSearchText);
                var pattern = $@"\b{escapedSearchText}\b";
                var regex = new Regex(pattern, RegexOptions.IgnoreCase);

                // Find matches and replace while preserving capitalization context
                var result = regex.Replace(trimmedSourceText, match =>
                {
                    return PreserveCapitalization(match.Value, replacementText);
                });

                // Preserve any trailing whitespace from original source text
                if (sourceText.Length > trimmedSourceText.Length)
                {
                    var trailingWhitespace = sourceText.Substring(trimmedSourceText.Length);
                    result += trailingWhitespace;
                }

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in text replacement with capitalization preservation. SearchText: {SearchText}, ReplacementText: {ReplacementText}", searchText, replacementText);
                return sourceText; // Return original text on error
            }
        }

        private string PreserveCapitalization(string originalMatch, string replacementText)
        {
            try
            {
                if (string.IsNullOrEmpty(originalMatch) || string.IsNullOrEmpty(replacementText))
                    return replacementText;

                var result = new StringBuilder(replacementText.Length);

                // Analyze capitalization pattern of original match
                var isAllUpper = originalMatch.All(c => !char.IsLetter(c) || char.IsUpper(c));
                var isAllLower = originalMatch.All(c => !char.IsLetter(c) || char.IsLower(c));
                var isFirstUpper = char.IsUpper(originalMatch.FirstOrDefault(char.IsLetter));

                if (isAllUpper && originalMatch.Any(char.IsLetter))
                {
                    // Original is all uppercase - make replacement all uppercase
                    return replacementText.ToUpperInvariant();
                }
                else if (isAllLower && originalMatch.Any(char.IsLetter))
                {
                    // Original is all lowercase - make replacement all lowercase
                    return replacementText.ToLowerInvariant();
                }
                else if (isFirstUpper)
                {
                    // Original has first letter capitalized - capitalize first letter of replacement
                    return CapitalizeFirstLetter(replacementText);
                }
                else
                {
                    // Default: return replacement text as provided by user
                    return replacementText;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Error preserving capitalization for match '{OriginalMatch}' -> '{ReplacementText}': {Error}", originalMatch, replacementText, ex.Message);
                return replacementText; // Fallback to user's original capitalization
            }
        }

        private string CapitalizeFirstLetter(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            var firstLetterIndex = text.ToList().FindIndex(char.IsLetter);
            if (firstLetterIndex == -1)
                return text;

            var chars = text.ToCharArray();
            chars[firstLetterIndex] = char.ToUpperInvariant(chars[firstLetterIndex]);

            return new string(chars);
        }
    }
}