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

        /// <summary>
        /// NEW METHOD: Processes text replacements using an already opened WordprocessingDocument to prevent corruption
        /// </summary>
        public Task<int> ProcessTextReplacementsInSessionAsync(WordprocessingDocument wordDocument, CoreDocument document, IEnumerable<TextReplacementRule> rules, CancellationToken cancellationToken = default)
        {
            try
            {
                var activeRules = rules.Where(r => r.IsEnabled && !string.IsNullOrWhiteSpace(r.SourceText) && !string.IsNullOrWhiteSpace(r.ReplacementText)).ToList();

                if (!activeRules.Any())
                {
                    _logger.LogDebug("No active text replacement rules found for document: {FileName}", document.FileName);
                    return Task.FromResult(0);
                }

                _logger.LogInformation("Processing {Count} text replacement rules in session for document: {FileName}", activeRules.Count, document.FileName);

                var mainPart = wordDocument.MainDocumentPart;
                if (mainPart?.Document?.Body == null)
                {
                    _logger.LogWarning("No document body found for text replacement: {FileName}", document.FileName);
                    return Task.FromResult(0);
                }

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

                        _logger.LogDebug("Text replacement in session element: '{OriginalText}' -> '{NewText}'", originalText, modifiedText);
                    }
                }

                _logger.LogInformation("Text replacement processing completed in session for document: {FileName}, replacements made: {Count}", document.FileName, totalReplacements);
                return Task.FromResult(totalReplacements);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing text replacements in session for document: {FileName}", document.FileName);
                throw;
            }
        }

        /// <summary>
        /// LEGACY METHOD: Opens document independently - can cause corruption
        /// </summary>
        [System.Obsolete("Use ProcessTextReplacementsInSessionAsync to prevent file corruption")]
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
                if (string.IsNullOrEmpty(sourceText) || string.IsNullOrWhiteSpace(searchText))
                    return sourceText;

                // Use a more robust regex that handles various word boundary scenarios
                var escapedSearchText = Regex.Escape(searchText);
                var pattern = $@"(?i)\b{escapedSearchText}\b";

                var result = Regex.Replace(sourceText, pattern, match =>
                {
                    // Preserve capitalization based on the original match
                    return PreserveCapitalization(match.Value, replacementText);
                });

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