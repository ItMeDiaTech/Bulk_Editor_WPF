using BulkEditor.Core.Configuration;
using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using BulkEditor.Infrastructure.Utilities;
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
        public Task<int> ProcessTextReplacementsInSessionAsync(WordprocessingDocument wordDocument, CoreDocument document, IEnumerable<TextReplacementRule> rules, bool trackChanges = false, CancellationToken cancellationToken = default)
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

                    // CRITICAL FIX: Handle complex paragraphs (with hyperlinks) differently
                    if (HasComplexElements(paragraph))
                    {
                        // For complex paragraphs, apply replacement rules directly to individual text elements
                        var originalText = GetParagraphText(paragraph);
                        
                        if (!string.IsNullOrEmpty(originalText))
                        {
                            // Check if any rules would apply to this paragraph's text
                            var wouldApply = activeRules.Any(rule => 
                                ReplaceTextExact(originalText, rule.SourceText, rule.ReplacementText) != originalText);
                            
                            if (wouldApply)
                            {
                                UpdateComplexParagraphText(paragraph, activeRules, trackChanges);
                                totalReplacements++; // Count as one replacement for the paragraph
                                
                                // Log the change in document
                                document.ChangeLog.Changes.Add(new ChangeEntry
                                {
                                    Type = ChangeType.TextReplaced,
                                    Description = "Text replaced in complex paragraph (element-level)",
                                    OldValue = originalText,
                                    NewValue = GetParagraphText(paragraph), // Get updated text
                                    ElementId = Guid.NewGuid().ToString(),
                                    Details = "Applied replacement rules to individual text elements"
                                });

                                _logger.LogDebug("Text replacement in complex paragraph: '{OriginalText}'", originalText);
                            }
                        }
                    }
                    else
                    {
                        // Handle simple paragraphs with the existing consolidated approach
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
                            // CRITICAL FIX: Validate paragraph structure before modification
                            if (ValidateParagraphForTextReplacement(paragraph))
                            {
                                UpdateParagraphText(paragraph, modifiedText, trackChanges);
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
                            else
                            {
                                _logger.LogWarning("Skipped text replacement for paragraph due to validation failure");
                            }
                        }
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
                            UpdateParagraphText(paragraph, modifiedText, false);
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

                // CRITICAL FIX: Handle multi-word phrases correctly by using lookbehind/lookahead instead of word boundaries
                // Word boundaries (\b) don't work properly with multi-word phrases containing spaces
                var escapedSearchText = Regex.Escape(searchText);
                
                // Use negative lookbehind/lookahead to ensure we match whole words/phrases, not partial matches
                // (?<!\w) = not preceded by a word character
                // (?!\w) = not followed by a word character  
                // This handles both single words and multi-word phrases correctly
                var pattern = $@"(?i)(?<!\w){escapedSearchText}(?!\w)";

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
        /// Supports track changes using run-level processing for proper deletion/insertion tracking
        /// </summary>
        private void UpdateParagraphText(Paragraph paragraph, string newText, bool trackChanges = false)
        {
            try
            {
                if (trackChanges)
                {
                    // Use run-level track changes for proper deletion/insertion tracking
                    UpdateParagraphTextWithTrackChanges(paragraph, newText);
                }
                else
                {
                    // Use standard non-tracked update
                    UpdateParagraphTextInternal(paragraph, newText);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating paragraph text. Falling back to simple text update.");
                
                // Fallback: update first simple text element found
                var firstText = paragraph.Descendants<Text>()
                    .FirstOrDefault(t => IsSimpleTextElement(t));
                
                if (firstText != null)
                {
                    firstText.Text = newText;
                }
                else
                {
                    _logger.LogWarning("No simple text elements found for fallback text replacement");
                }
            }
        }

        /// <summary>
        /// Updates paragraph text with track changes using run-level processing
        /// Creates proper deletion and insertion tracking for Word visibility
        /// </summary>
        private void UpdateParagraphTextWithTrackChanges(Paragraph paragraph, string newText)
        {
            var oldText = GetParagraphText(paragraph);
            
            if (oldText == newText)
            {
                _logger.LogDebug("No text changes detected, skipping track changes processing");
                return;
            }
            
            // For simple paragraphs, apply track changes at run level
            var runs = paragraph.Elements<Run>().ToList();
            if (!runs.Any())
            {
                // No runs exist, create a simple tracked insertion
                var insertedRun = OpenXmlHelper.CreateTrackedInsertion(newText);
                paragraph.Append(insertedRun);
                _logger.LogDebug("Created tracked insertion for new text in empty paragraph");
                return;
            }

            // Process each run to apply track changes using the existing OpenXMLHelper method
            var textChanged = false;
            foreach (var run in runs.ToList()) // ToList to avoid modification during enumeration
            {
                var runText = string.Join("", run.Descendants<Text>().Select(t => t.Text ?? ""));
                
                if (string.IsNullOrEmpty(runText))
                    continue;

                // Check if this run's text needs to be changed
                // We'll do a simple approach: if the run contains text that should be different in newText,
                // mark the whole run as deleted and create a new insertion
                if (oldText.Contains(runText) && !newText.Contains(runText))
                {
                    // This run's text is being removed or changed
                    var deletedRun = OpenXmlHelper.CreateTrackedDeletion(run);
                    paragraph.InsertBefore(deletedRun, run);
                    run.Remove();
                    textChanged = true;
                    _logger.LogDebug("Created tracked deletion for run text: '{RunText}'", runText);
                }
            }

            // If we made deletions, add the new text as an insertion
            if (textChanged && !string.IsNullOrEmpty(newText))
            {
                // Find the first run to get formatting
                var firstRun = paragraph.Elements<Run>().FirstOrDefault();
                var formatting = firstRun?.GetFirstChild<RunProperties>();
                
                var insertedRun = OpenXmlHelper.CreateTrackedInsertion(newText, formatting);
                
                // Insert at the beginning of the paragraph (after any deleted runs)
                var lastDeletedRun = paragraph.Elements<DeletedRun>().LastOrDefault();
                if (lastDeletedRun != null)
                {
                    paragraph.InsertAfter(insertedRun, lastDeletedRun);
                }
                else
                {
                    paragraph.PrependChild(insertedRun);
                }
                
                _logger.LogDebug("Created tracked insertion for new text: '{NewText}'", newText);
            }
        }

        /// <summary>
        /// Internal method for updating paragraph text without track changes
        /// CRITICAL FIX: Uses text-only updates to preserve document structure and prevent corruption
        /// </summary>
        private void UpdateParagraphTextInternal(Paragraph paragraph, string newText)
        {
            // CRITICAL FIX: Only update text content, preserve all runs and structure
            // This prevents "Word found unreadable content" errors by maintaining document integrity
            
            // Check if paragraph contains complex elements (hyperlinks, fields, etc.)
            if (HasComplexElements(paragraph))
            {
                // For complex paragraphs, we need to apply replacements differently
                // This method is called after paragraph-level processing, so we can't apply rules here
                // Log a warning that complex paragraph text replacement needs special handling
                _logger.LogDebug("Complex paragraph detected - text replacement should be handled at element level");
                return;
            }

            // For simple paragraphs, consolidate all text content into the first run
            var runs = paragraph.Elements<Run>().ToList();
            if (!runs.Any())
            {
                // No runs exist, create a simple one
                var newRun = new Run();
                newRun.Append(new Text(newText));
                paragraph.Append(newRun);
                return;
            }

            // Use the first run and preserve its formatting
            var firstRun = runs.First();
            var runProperties = firstRun.GetFirstChild<RunProperties>();

            // Clear text from first run and add new text
            var textElements = firstRun.Elements<Text>().ToList();
            foreach (var textElement in textElements)
            {
                textElement.Remove();
            }
            firstRun.Append(new Text(newText));

            // Remove additional runs (but preserve the first one with its formatting)
            for (int i = 1; i < runs.Count; i++)
            {
                // Only remove runs that contain only text (not complex elements)
                if (IsSimpleTextRun(runs[i]))
                {
                    runs[i].Remove();
                }
            }
        }

        /// <summary>
        /// Checks if a paragraph contains complex elements that need careful handling
        /// CRITICAL FIX: Include hyperlinks for proper structure preservation
        /// </summary>
        private bool HasComplexElements(Paragraph paragraph)
        {
            // Check for elements that need careful text replacement to preserve structure
            return paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Hyperlink>().Any() ||
                   paragraph.Descendants<FieldCode>().Any() ||
                   paragraph.Descendants<Drawing>().Any() ||
                   paragraph.Descendants<FieldChar>().Count() > 2; // Complex field structures
        }

        /// <summary>
        /// Updates text elements individually while preserving complex structure like hyperlinks
        /// Supports track changes for individual text elements using OpenXMLHelper
        /// </summary>
        private void UpdateComplexParagraphText(Paragraph paragraph, IEnumerable<TextReplacementRule> rules, bool trackChanges = false)
        {
            // Find all text elements that can be safely modified
            var textElements = paragraph.Descendants<Text>()
                .Where(t => IsSimpleTextElement(t) && !string.IsNullOrWhiteSpace(t.Text))
                .ToList();

            if (!textElements.Any())
            {
                _logger.LogDebug("No simple text elements found for replacement in complex paragraph");
                return;
            }

            // Apply replacement rules to each individual text element
            var replacementsMade = 0;
            foreach (var textElement in textElements)
            {
                var currentText = textElement.Text ?? string.Empty;
                var originalText = currentText;

                // Apply each replacement rule to this text element
                foreach (var rule in rules)
                {
                    if (trackChanges && currentText.Contains(rule.SourceText))
                    {
                        // Use OpenXMLHelper for run-level track changes
                        var parentRun = textElement.Ancestors<Run>().FirstOrDefault();
                        if (parentRun != null)
                        {
                            OpenXmlHelper.ReplaceTextInRun(parentRun, rule.SourceText, rule.ReplacementText, trackChanges);
                            replacementsMade++;
                            _logger.LogDebug("Applied tracked text replacement in complex paragraph run: '{OldText}' -> '{NewText}'", rule.SourceText, rule.ReplacementText);
                        }
                        else
                        {
                            // Fallback to direct text replacement if no parent run found
                            currentText = ReplaceTextExact(currentText, rule.SourceText, rule.ReplacementText);
                        }
                    }
                    else
                    {
                        // Standard text replacement without track changes
                        currentText = ReplaceTextExact(currentText, rule.SourceText, rule.ReplacementText);
                    }
                }

                // Update the text element if changes were made (for non-tracked changes)
                if (currentText != originalText && !trackChanges)
                {
                    textElement.Text = currentText;
                    replacementsMade++;
                }
            }

            if (replacementsMade > 0)
            {
                var trackingStatus = trackChanges ? "with track changes" : "without track changes";
                _logger.LogDebug("Applied text replacement to {Count} elements in complex paragraph {TrackingStatus}", replacementsMade, trackingStatus);
            }
        }

        /// <summary>
        /// Checks if a run contains only simple text (no complex elements)
        /// </summary>
        private bool IsSimpleTextRun(Run run)
        {
            return !run.Descendants<FieldCode>().Any() &&
                   !run.Descendants<FieldChar>().Any() &&
                   !run.Descendants<Drawing>().Any() &&
                   run.Elements().All(e => e is RunProperties || e is Text);
        }

        /// <summary>
        /// Checks if a text element is simple (not part of dangerous structures)
        /// CRITICAL FIX: Allow hyperlink text and normal formatted text
        /// </summary>
        private bool IsSimpleTextElement(Text textElement)
        {
            // Only exclude text elements that are inside dangerous field codes
            // Allow hyperlink display text, table text, and normal formatted text
            return !textElement.Ancestors<FieldCode>().Any() &&
                   textElement.Ancestors<Run>().Any(); // Must be in a run
        }

        /// <summary>
        /// Validates if a paragraph is safe for text replacement without causing document corruption
        /// CRITICAL FIX: Simplified validation - only blocks genuinely dangerous structures
        /// </summary>
        private bool ValidateParagraphForTextReplacement(Paragraph paragraph)
        {
            try
            {
                // Only block truly dangerous structures that can corrupt documents
                // Allow normal table text, hyperlink text, and formatted text
                
                // Check for genuinely problematic structures
                var hasFieldCodes = paragraph.Descendants<FieldCode>().Any();
                var hasDrawingElements = paragraph.Descendants<Drawing>().Any();
                var hasComplexFields = paragraph.Descendants<FieldChar>().Count() > 2; // Complex field structure
                
                // Only block paragraphs with dangerous field codes or complex graphics
                if (hasFieldCodes || hasDrawingElements || hasComplexFields)
                {
                    _logger.LogDebug("Paragraph contains field codes, drawings, or complex fields - skipping text replacement for safety");
                    return false;
                }

                // Check if paragraph has any text content at all
                var hasText = paragraph.Descendants<Text>()
                    .Any(t => !string.IsNullOrWhiteSpace(t.Text));

                if (!hasText)
                {
                    _logger.LogDebug("Paragraph contains no text content for replacement");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error validating paragraph for text replacement, skipping for safety");
                return false;
            }
        }
    }
}