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

                // CRITICAL FIX: Process paragraphs instead of individual text elements
                // This handles text that is fragmented across multiple Text elements due to formatting
                var paragraphs = mainPart.Document.Body.Descendants<Paragraph>().ToList();

                foreach (var paragraph in paragraphs)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // Get all text content from the paragraph (consolidated)
                    var paragraphText = GetParagraphText(paragraph);
                    
                    if (string.IsNullOrEmpty(paragraphText))
                        continue;

                    var originalText = paragraphText;
                    var modifiedText = originalText;
                    var replacementsMadeInParagraph = 0;

                    foreach (var rule in activeRules)
                    {
                        // Use exact text replacement (case-insensitive find, exact replace)
                        var replacementResult = ReplaceTextExact(
                            modifiedText, rule.SourceText, rule.ReplacementText);

                        if (replacementResult != modifiedText)
                        {
                            modifiedText = replacementResult;
                            replacementsMadeInParagraph++;
                        }
                    }

                    // Update paragraph if any replacements were made
                    if (replacementsMadeInParagraph > 0)
                    {
                        UpdateParagraphText(paragraph, modifiedText);
                        totalReplacements += replacementsMadeInParagraph;

                        // Log the change in document
                        document.ChangeLog.Changes.Add(new ChangeEntry
                        {
                            Type = ChangeType.TextReplaced,
                            Description = "Text replaced using replacement rules (paragraph-level)",
                            OldValue = originalText,
                            NewValue = modifiedText,
                            ElementId = Guid.NewGuid().ToString(),
                            Details = $"Applied {replacementsMadeInParagraph} replacement rule(s)"
                        });

                        _logger.LogDebug("Text replacement in session paragraph: '{OriginalText}' -> '{NewText}'", originalText, modifiedText);
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

                    // CRITICAL FIX: Process paragraphs instead of individual text elements (legacy method)
                    // This handles text that is fragmented across multiple Text elements due to formatting
                    var paragraphs = mainPart.Document.Body.Descendants<Paragraph>().ToList();

                    foreach (var paragraph in paragraphs)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        // Get all text content from the paragraph (consolidated)
                        var paragraphText = GetParagraphText(paragraph);
                        
                        if (string.IsNullOrEmpty(paragraphText))
                            continue;

                        var originalText = paragraphText;
                        var modifiedText = originalText;
                        var replacementsMadeInParagraph = 0;

                        foreach (var rule in activeRules)
                        {
                            // Use exact text replacement (case-insensitive find, exact replace)
                            var replacementResult = ReplaceTextExact(
                                modifiedText, rule.SourceText, rule.ReplacementText);

                            if (replacementResult != modifiedText)
                            {
                                modifiedText = replacementResult;
                                replacementsMadeInParagraph++;
                            }
                        }

                        // Update paragraph if any replacements were made
                        if (replacementsMadeInParagraph > 0)
                        {
                            UpdateParagraphText(paragraph, modifiedText);
                            totalReplacements += replacementsMadeInParagraph;

                            // Log the change in document
                            document.ChangeLog.Changes.Add(new ChangeEntry
                            {
                                Type = ChangeType.TextReplaced,
                                Description = "Text replaced using replacement rules (paragraph-level, legacy)",
                                OldValue = originalText,
                                NewValue = modifiedText,
                                ElementId = Guid.NewGuid().ToString(),
                                Details = $"Applied {replacementsMadeInParagraph} replacement rule(s)"
                            });

                            _logger.LogDebug("Text replacement in legacy paragraph: '{OriginalText}' -> '{NewText}'", originalText, modifiedText);
                        }
                    }

                    if (totalReplacements > 0)
                    {
                        // CRITICAL FIX: Do NOT save here - conflicts with session-based saves
                        // The document will be saved by the main session processor
                        // mainPart.Document.Save(); // REMOVED - causes save conflicts
                        _logger.LogInformation("Processed {Count} text replacements in legacy method: {FileName} (save handled by session)", totalReplacements, document.FileName);
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

        /// <summary>
        /// Replaces text with exact replacement text (case-insensitive find, exact replace)
        /// This is the preferred method that respects user's exact capitalization preferences
        /// </summary>
        public string ReplaceTextExact(string sourceText, string searchText, string replacementText)
        {
            try
            {
                if (string.IsNullOrEmpty(sourceText) || string.IsNullOrWhiteSpace(searchText))
                    return sourceText;

                // Use regex with case-insensitive matching and word boundaries
                var escapedSearchText = Regex.Escape(searchText);
                var pattern = $@"(?i)\b{escapedSearchText}\b";

                // Replace with exact text provided by user (no capitalization changes)
                var result = Regex.Replace(sourceText, pattern, replacementText);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in exact text replacement. SearchText: {SearchText}, ReplacementText: {ReplacementText}", searchText, replacementText);
                return sourceText; // Return original text on error
            }
        }

        /// <summary>
        /// Gets all text content from a paragraph, consolidating across all runs and text elements
        /// </summary>
        private string GetParagraphText(Paragraph paragraph)
        {
            var textBuilder = new StringBuilder();
            
            foreach (var run in paragraph.Descendants<Run>())
            {
                foreach (var text in run.Descendants<Text>())
                {
                    textBuilder.Append(text.Text ?? string.Empty);
                }
            }
            
            return textBuilder.ToString();
        }

        /// <summary>
        /// Updates paragraph text while preserving formatting structure
        /// Uses consolidation approach - creates a single run with the new text
        /// </summary>
        private void UpdateParagraphText(Paragraph paragraph, string newText)
        {
            try
            {
                // Get the first run to preserve its properties
                var firstRun = paragraph.Descendants<Run>().FirstOrDefault();
                var runProperties = firstRun?.GetFirstChild<RunProperties>()?.CloneNode(true);

                // Remove all existing runs
                var runsToRemove = paragraph.Descendants<Run>().ToList();
                foreach (var run in runsToRemove)
                {
                    run.Remove();
                }

                // Create a new run with the updated text
                var newRun = new Run();
                
                // Apply preserved formatting if available
                if (runProperties != null)
                {
                    newRun.Append(runProperties);
                }

                // Add the new text
                newRun.Append(new Text(newText));

                // Add the new run to the paragraph
                paragraph.Append(newRun);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating paragraph text. Falling back to simple text update.");
                
                // Fallback: update first text element found
                var firstText = paragraph.Descendants<Text>().FirstOrDefault();
                if (firstText != null)
                {
                    firstText.Text = newText;
                }
            }
        }
    }
}