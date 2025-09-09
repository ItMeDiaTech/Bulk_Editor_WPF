using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using OpenXmlDocument = DocumentFormat.OpenXml.Wordprocessing.Document;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Implementation of text optimization service using OpenXML
    /// </summary>
    public class TextOptimizer : ITextOptimizer
    {
        private readonly ILoggingService _logger;
        private readonly TextOptimizationSettings _settings;

        public TextOptimizer(ILoggingService logger, TextOptimizationSettings? settings = null)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _settings = settings ?? new TextOptimizationSettings();
        }

        /// <summary>
        /// NEW METHOD: Optimizes text in an already opened WordprocessingDocument to prevent corruption
        /// </summary>
        public async Task<int> OptimizeDocumentTextInSessionAsync(WordprocessingDocument wordDocument, BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken = default)
        {
            try
            {
                _logger.LogInformation("Starting text optimization in session for document: {FileName}", document.FileName);

                var mainPart = wordDocument.MainDocumentPart;
                if (mainPart?.Document?.Body == null)
                {
                    _logger.LogWarning("No document body found for text optimization: {FileName}", document.FileName);
                    return 0;
                }

                var changesMade = 0;

                // 1. Optimize whitespace
                if (_settings.RemoveExtraSpaces)
                {
                    changesMade += await OptimizeWhitespaceInDocumentAsync(mainPart, document, cancellationToken);
                }

                // 2. Remove empty paragraphs
                if (_settings.RemoveEmptyParagraphs)
                {
                    changesMade += await RemoveEmptyParagraphsAsync(mainPart, document, cancellationToken);
                }

                // 3. Standardize line breaks
                if (_settings.StandardizeLineBreaks)
                {
                    changesMade += await StandardizeLineBreaksAsync(mainPart, document, cancellationToken);
                }

                // 4. Optimize table formatting
                if (_settings.OptimizeTableFormatting)
                {
                    changesMade += await OptimizeTableFormattingAsync(mainPart, document, cancellationToken);
                }

                // 5. Optimize list formatting
                if (_settings.OptimizeListFormatting)
                {
                    changesMade += await OptimizeListFormattingAsync(mainPart, document, cancellationToken);
                }

                // Log summary change
                if (changesMade > 0)
                {
                    document.ChangeLog.Changes.Add(new ChangeEntry
                    {
                        Type = ChangeType.TextOptimized,
                        Description = $"Text optimization completed: {changesMade} improvements made",
                        Details = $"Optimized whitespace, paragraphs, formatting"
                    });

                    _logger.LogInformation("Text optimization completed in session for {FileName}: {Changes} improvements made",
                        document.FileName, changesMade);
                }
                else
                {
                    _logger.LogInformation("No text optimization needed for document: {FileName}", document.FileName);
                }

                return changesMade;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during text optimization in session for document: {FileName}", document.FileName);

                document.ProcessingErrors.Add(new ProcessingError { Message = $"Text optimization failed: {ex.Message}", Severity = ErrorSeverity.Error });

                document.ChangeLog.Changes.Add(new ChangeEntry
                {
                    Type = ChangeType.Error,
                    Description = "Text optimization failed",
                    Details = ex.Message
                });

                throw;
            }
        }

        /// <summary>
        /// LEGACY METHOD: Opens document independently - can cause corruption
        /// </summary>
        [System.Obsolete("Use OptimizeDocumentTextInSessionAsync to prevent file corruption")]
        public async Task<BulkEditor.Core.Entities.Document> OptimizeDocumentTextAsync(BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken = default)
        {
            try
            {
                _logger.LogInformation("Starting text optimization for document: {FileName}", document.FileName);

                using var wordDocument = WordprocessingDocument.Open(document.FilePath, true);
                var mainPart = wordDocument.MainDocumentPart;

                if (mainPart?.Document?.Body != null)
                {
                    var changesMade = 0;

                    // 1. Optimize whitespace
                    if (_settings.RemoveExtraSpaces)
                    {
                        changesMade += await OptimizeWhitespaceInDocumentAsync(mainPart, document, cancellationToken);
                    }

                    // 2. Remove empty paragraphs
                    if (_settings.RemoveEmptyParagraphs)
                    {
                        changesMade += await RemoveEmptyParagraphsAsync(mainPart, document, cancellationToken);
                    }

                    // 3. Standardize line breaks
                    if (_settings.StandardizeLineBreaks)
                    {
                        changesMade += await StandardizeLineBreaksAsync(mainPart, document, cancellationToken);
                    }

                    // 4. Optimize table formatting
                    if (_settings.OptimizeTableFormatting)
                    {
                        changesMade += await OptimizeTableFormattingAsync(mainPart, document, cancellationToken);
                    }

                    // 5. Optimize list formatting
                    if (_settings.OptimizeListFormatting)
                    {
                        changesMade += await OptimizeListFormattingAsync(mainPart, document, cancellationToken);
                    }

                    // CRITICAL FIX: Changes tracked but save handled by main session
                    if (changesMade > 0)
                    {
                        // REMOVED: mainPart.Document.Save(); - conflicts with session-based saves
                        // The document will be saved by the main session processor

                        // Log summary change
                        document.ChangeLog.Changes.Add(new ChangeEntry
                        {
                            Type = ChangeType.TextOptimized,
                            Description = $"Text optimization completed: {changesMade} improvements made",
                            Details = $"Optimized whitespace, paragraphs, formatting"
                        });

                        _logger.LogInformation("Text optimization completed for {FileName}: {Changes} improvements made",
                            document.FileName, changesMade);
                    }
                    else
                    {
                        _logger.LogInformation("No text optimization needed for document: {FileName}", document.FileName);
                    }
                }

                return document;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during text optimization for document: {FileName}", document.FileName);

                document.ProcessingErrors.Add(new ProcessingError { Message = $"Text optimization failed: {ex.Message}", Severity = ErrorSeverity.Error });

                document.ChangeLog.Changes.Add(new ChangeEntry
                {
                    Type = ChangeType.Error,
                    Description = "Text optimization failed",
                    Details = ex.Message
                });

                return document;
            }
        }

        public async Task<string> OptimizeWhitespaceAsync(string text, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            try
            {
                // Only remove multiple consecutive spaces (2 or more) and replace with single space
                // This preserves necessary single spaces after colons, semicolons, periods, etc.
                var optimized = Regex.Replace(text, @" {2,}", " ");

                // Standardize line endings
                optimized = Regex.Replace(optimized, @"\r\n|\r|\n", Environment.NewLine);

                // Limit consecutive line breaks
                var maxBreaks = new string('\n', _settings.MaxConsecutiveLineBreaks);
                var breakPattern = @"\n{" + (_settings.MaxConsecutiveLineBreaks + 1) + ",}";
                optimized = Regex.Replace(optimized, breakPattern, maxBreaks);

                await Task.CompletedTask; // Make method async
                return optimized;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error optimizing whitespace in text");
                return text; // Return original text if optimization fails
            }
        }

        public async Task<BulkEditor.Core.Entities.Document> StandardizeParagraphFormattingAsync(BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken = default)
        {
            try
            {
                using var wordDocument = WordprocessingDocument.Open(document.FilePath, true);
                var mainPart = wordDocument.MainDocumentPart;

                if (mainPart?.Document?.Body != null)
                {
                    var paragraphs = mainPart.Document.Body.Elements<Paragraph>().ToList();
                    var changesMade = 0;

                    foreach (var paragraph in paragraphs)
                    {
                        if (StandardizeParagraphSpacing(paragraph))
                        {
                            changesMade++;
                        }
                    }

                    if (changesMade > 0)
                    {
                        // REMOVED: mainPart.Document.Save(); - conflicts with session-based saves
                        // The document will be saved by the main session processor

                        document.ChangeLog.Changes.Add(new ChangeEntry
                        {
                            Type = ChangeType.TextOptimized,
                            Description = $"Standardized formatting for {changesMade} paragraphs"
                        });
                    }
                }

                await Task.CompletedTask;
                return document;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error standardizing paragraph formatting");
                throw;
            }
        }

        public async Task<BulkEditor.Core.Entities.Document> RemoveUnnecessaryFormattingAsync(BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken = default)
        {
            try
            {
                using var wordDocument = WordprocessingDocument.Open(document.FilePath, true);
                var mainPart = wordDocument.MainDocumentPart;

                if (mainPart?.Document?.Body != null)
                {
                    var changesMade = 0;

                    // Remove unnecessary run properties
                    var runs = mainPart.Document.Body.Descendants<Run>().ToList();
                    foreach (var run in runs)
                    {
                        if (CleanRunFormatting(run))
                        {
                            changesMade++;
                        }
                    }

                    if (changesMade > 0)
                    {
                        // REMOVED: mainPart.Document.Save(); - conflicts with session-based saves
                        // The document will be saved by the main session processor

                        document.ChangeLog.Changes.Add(new ChangeEntry
                        {
                            Type = ChangeType.TextOptimized,
                            Description = $"Removed unnecessary formatting from {changesMade} text runs"
                        });
                    }
                }

                await Task.CompletedTask;
                return document;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error removing unnecessary formatting");
                throw;
            }
        }

        public async Task<BulkEditor.Core.Entities.Document> OptimizeDocumentStructureAsync(BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken = default)
        {
            try
            {
                using var wordDocument = WordprocessingDocument.Open(document.FilePath, true);
                var mainPart = wordDocument.MainDocumentPart;

                if (mainPart?.Document?.Body != null)
                {
                    var changesMade = 0;

                    // Optimize heading hierarchy
                    changesMade += OptimizeHeadingHierarchy(mainPart.Document.Body);

                    // Optimize section breaks
                    changesMade += OptimizeSectionBreaks(mainPart.Document.Body);

                    if (changesMade > 0)
                    {
                        // REMOVED: mainPart.Document.Save(); - conflicts with session-based saves
                        // The document will be saved by the main session processor

                        document.ChangeLog.Changes.Add(new ChangeEntry
                        {
                            Type = ChangeType.TextOptimized,
                            Description = $"Optimized document structure: {changesMade} improvements"
                        });
                    }
                }

                await Task.CompletedTask;
                return document;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error optimizing document structure");
                throw;
            }
        }

        #region Private Helper Methods

        private async Task<int> OptimizeWhitespaceInDocumentAsync(MainDocumentPart mainPart, BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken)
        {
            var changesMade = 0;
            var textElements = mainPart.Document.Body?.Descendants<Text>().ToList() ?? new List<Text>();

            foreach (var textElement in textElements)
            {
                if (!string.IsNullOrEmpty(textElement.Text))
                {
                    var originalText = textElement.Text;
                    var optimizedText = await OptimizeWhitespaceAsync(originalText, cancellationToken);

                    if (originalText != optimizedText)
                    {
                        textElement.Text = optimizedText;
                        changesMade++;
                    }
                }
            }

            return changesMade;
        }

        private async Task<int> RemoveEmptyParagraphsAsync(MainDocumentPart mainPart, BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken)
        {
            var paragraphs = mainPart.Document.Body?.Elements<Paragraph>().ToList() ?? new List<Paragraph>();
            var removedCount = 0;

            foreach (var paragraph in paragraphs)
            {
                if (IsEmptyParagraph(paragraph))
                {
                    paragraph.Remove();
                    removedCount++;
                }
            }

            await Task.CompletedTask;
            return removedCount;
        }

        private async Task<int> StandardizeLineBreaksAsync(MainDocumentPart mainPart, BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken)
        {
            var changesMade = 0;
            var breaks = mainPart.Document.Body?.Descendants<Break>().ToList() ?? new List<Break>();

            // Remove excessive line breaks
            var consecutiveBreaks = 0;
            Break? previousBreak = null;

            foreach (var lineBreak in breaks)
            {
                if (previousBreak != null && AreConsecutiveBreaks(previousBreak, lineBreak))
                {
                    consecutiveBreaks++;
                    if (consecutiveBreaks >= _settings.MaxConsecutiveLineBreaks)
                    {
                        lineBreak.Remove();
                        changesMade++;
                        continue;
                    }
                }
                else
                {
                    consecutiveBreaks = 1;
                }

                previousBreak = lineBreak;
            }

            await Task.CompletedTask;
            return changesMade;
        }

        private async Task<int> OptimizeTableFormattingAsync(MainDocumentPart mainPart, BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken)
        {
            var tables = mainPart.Document.Body?.Descendants<Table>().ToList() ?? new List<Table>();
            var changesMade = 0;

            foreach (var table in tables)
            {
                // Remove empty table cells
                var emptyCells = table.Descendants<TableCell>()
                    .Where(cell => IsEmptyTableCell(cell))
                    .ToList();

                foreach (var emptyCell in emptyCells)
                {
                    // Add minimal content to prevent Word issues
                    if (!emptyCell.Elements<Paragraph>().Any())
                    {
                        emptyCell.AppendChild(new Paragraph(new Run(new Text(""))));
                        changesMade++;
                    }
                }
            }

            await Task.CompletedTask;
            return changesMade;
        }

        private async Task<int> OptimizeListFormattingAsync(MainDocumentPart mainPart, BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken)
        {
            var changesMade = 0;
            var paragraphs = mainPart.Document.Body?.Elements<Paragraph>().ToList() ?? new List<Paragraph>();

            foreach (var paragraph in paragraphs)
            {
                // Optimize numbered list formatting
                var numbering = paragraph.ParagraphProperties?.NumberingProperties;
                if (numbering != null)
                {
                    // Ensure consistent list formatting
                    if (StandardizeListFormatting(paragraph))
                    {
                        changesMade++;
                    }
                }
            }

            await Task.CompletedTask;
            return changesMade;
        }

        private bool IsEmptyParagraph(Paragraph paragraph)
        {
            // Check if paragraph has no text content
            var text = paragraph.InnerText?.Trim();
            return string.IsNullOrEmpty(text);
        }

        private bool IsEmptyTableCell(TableCell cell)
        {
            var text = cell.InnerText?.Trim();
            return string.IsNullOrEmpty(text);
        }

        private bool AreConsecutiveBreaks(Break break1, Break break2)
        {
            // Simple heuristic - in real implementation would check document position
            return break1.Parent == break2.Parent;
        }

        private bool StandardizeParagraphSpacing(Paragraph paragraph)
        {
            var properties = paragraph.ParagraphProperties;
            if (properties == null)
            {
                properties = new ParagraphProperties();
                paragraph.PrependChild(properties);
            }

            var spacing = properties.SpacingBetweenLines;
            bool changed = false;

            // Standardize spacing (e.g., single spacing)
            if (spacing == null)
            {
                spacing = new SpacingBetweenLines();
                properties.AppendChild(spacing);
                changed = true;
            }

            // Set standard line spacing
            if (spacing.Line?.Value != "240") // 12pt = 240 twentieths of a point
            {
                spacing.Line = "240";
                spacing.LineRule = LineSpacingRuleValues.Auto;
                changed = true;
            }

            return changed;
        }

        private bool CleanRunFormatting(Run run)
        {
            var properties = run.RunProperties;
            if (properties == null)
                return false;

            bool changed = false;

            // Remove redundant formatting that matches document defaults
            if (properties.FontSize?.Val?.Value == "22") // Default 11pt = 22 half-points
            {
                properties.FontSize = null;
                changed = true;
            }

            // Remove empty properties
            if (!properties.HasChildren)
            {
                run.RunProperties = null;
                changed = true;
            }

            return changed;
        }

        private int OptimizeHeadingHierarchy(Body body)
        {
            var changesMade = 0;
            var paragraphs = body.Elements<Paragraph>().ToList();

            foreach (var paragraph in paragraphs)
            {
                var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

                if (!string.IsNullOrEmpty(style) && style.StartsWith("Heading"))
                {
                    // Ensure proper heading hierarchy (this is a simplified implementation)
                    var properties = paragraph.ParagraphProperties;
                    if (properties != null)
                    {
                        // Add outline level if missing
                        if (properties.OutlineLevel == null)
                        {
                            var level = ExtractHeadingLevel(style);
                            if (level > 0)
                            {
                                properties.OutlineLevel = new OutlineLevel { Val = level - 1 };
                                changesMade++;
                            }
                        }
                    }
                }
            }

            return changesMade;
        }

        private int OptimizeSectionBreaks(Body body)
        {
            var changesMade = 0;
            var paragraphs = body.Elements<Paragraph>().ToList();

            // Remove unnecessary section breaks
            foreach (var paragraph in paragraphs)
            {
                var properties = paragraph.ParagraphProperties;
                var sectionProperties = properties?.SectionProperties;

                if (sectionProperties != null && IsUnnecessarySectionBreak(sectionProperties))
                {
                    if (properties != null)
                        properties.SectionProperties = null;
                    changesMade++;
                }
            }

            return changesMade;
        }

        private bool StandardizeListFormatting(Paragraph paragraph)
        {
            var properties = paragraph.ParagraphProperties;
            if (properties?.NumberingProperties == null)
                return false;

            // Ensure consistent indentation for list items
            var indentation = properties.Indentation;
            if (indentation == null)
            {
                indentation = new Indentation();
                properties.AppendChild(indentation);
            }

            // Set standard list indentation
            var numLevel = properties.NumberingProperties.NumberingLevelReference?.Val?.Value ?? 0;
            var expectedIndent = (numLevel * 720).ToString(); // 720 twentieths = 0.5 inch per level

            if (indentation.Left?.Value != expectedIndent)
            {
                indentation.Left = expectedIndent;
                return true;
            }

            return false;
        }

        private int ExtractHeadingLevel(string headingStyle)
        {
            // Extract numeric level from styles like "Heading1", "Heading2", etc.
            var match = Regex.Match(headingStyle, @"Heading(\d+)", RegexOptions.IgnoreCase);
            return match.Success && int.TryParse(match.Groups[1].Value, out var level) ? level : 1;
        }

        private bool IsUnnecessarySectionBreak(SectionProperties sectionProperties)
        {
            // Simple heuristic - section break is unnecessary if it doesn't change page layout
            var pageSize = sectionProperties.GetFirstChild<PageSize>();
            var pageMargin = sectionProperties.GetFirstChild<PageMargin>();

            // If section has default page settings, it might be unnecessary
            return pageSize == null && pageMargin == null;
        }

        #endregion
    }
}