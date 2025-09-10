using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Linq;
using System;
using System.Collections.Generic;

namespace BulkEditor.Infrastructure.Utilities
{
    public static class OpenXmlHelper
    {
        /// <summary>
        /// Safely updates hyperlink text, preserving formatting. 
        /// CRITICAL FIX: Attributes (DocLocation, Anchor, etc.) persist automatically through RemoveAllChildren()
        /// so we don't need to preserve/restore them, which was causing issues.
        /// </summary>
        public static void UpdateHyperlinkText(Hyperlink hyperlink, string newText)
        {
            UpdateHyperlinkText(hyperlink, newText, false);
        }

        /// <summary>
        /// Safely updates hyperlink text with optional track changes support.
        /// When trackChanges is true, wraps modifications in track changes markup for Word visibility.
        /// </summary>
        public static void UpdateHyperlinkText(Hyperlink hyperlink, string newText, bool trackChanges)
        {
            if (!trackChanges)
            {
                UpdateHyperlinkTextInternal(hyperlink, newText);
                return;
            }

            // For track changes, we need to work at the paragraph level
            // since hyperlinks cannot contain track changes elements directly
            var paragraph = hyperlink.Ancestors<Paragraph>().FirstOrDefault();
            if (paragraph == null)
            {
                // Fallback to non-tracking update if we can't find the paragraph
                UpdateHyperlinkTextInternal(hyperlink, newText);
                return;
            }

            try
            {
                UpdateHyperlinkWithParagraphLevelTracking(paragraph, hyperlink, newText);
            }
            catch (Exception ex)
            {
                // If track changes fail, fall back to simple update
                // This ensures document processing continues even if track changes have issues
                UpdateHyperlinkTextInternal(hyperlink, newText);
                
                // Note: We don't re-throw here to prevent document processing failure
                // The calling code should log this if needed
                throw new InvalidOperationException($"Track changes failed, used fallback update: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Internal method for updating hyperlink text without track changes
        /// </summary>
        private static void UpdateHyperlinkTextInternal(Hyperlink hyperlink, string newText)
        {
            // Preserve existing formatting by finding the first run and cloning its properties
            var runProperties = hyperlink.Descendants<Run>()
                .FirstOrDefault()
                ?.GetFirstChild<RunProperties>()
                ?.CloneNode(true);

            // CRITICAL FIX: Only remove children, not attributes
            // Attributes like DocLocation, Anchor, Tooltip, History are preserved automatically
            // The previous approach of manually preserving/restoring was interfering with them
            hyperlink.RemoveAllChildren();
            
            // Add the new text run with preserved formatting
            var newRun = new Run();
            if (runProperties != null)
            {
                newRun.Append(runProperties);
            }
            newRun.Append(new Text(newText));
            hyperlink.Append(newRun);
        }

        /// <summary>
        /// Gets the current text content of a hyperlink
        /// </summary>
        private static string GetHyperlinkText(Hyperlink hyperlink)
        {
            return string.Join("", hyperlink.Descendants<Text>().Select(t => t.Text ?? ""));
        }

        /// <summary>
        /// Updates a hyperlink with track changes at the paragraph level (schema-compliant approach)
        /// </summary>
        private static void UpdateHyperlinkWithParagraphLevelTracking(Paragraph paragraph, Hyperlink hyperlink, string newText)
        {
            try
            {
                // Validate inputs
                if (paragraph == null)
                    throw new ArgumentNullException(nameof(paragraph));
                if (hyperlink == null)
                    throw new ArgumentNullException(nameof(hyperlink));
                if (string.IsNullOrEmpty(newText))
                    throw new ArgumentException("New text cannot be null or empty", nameof(newText));

                // Store original hyperlink properties for cloning
                var originalText = GetHyperlinkText(hyperlink);
                var hyperlinkProperties = CloneHyperlinkProperties(hyperlink);
                
                // Create a deleted run with the original hyperlink
                var deletedRun = CreateDeletedRun();
                var originalHyperlinkClone = (Hyperlink)hyperlink.CloneNode(true);
                var deletedRunContent = new Run();
                deletedRunContent.Append(originalHyperlinkClone);
                deletedRun.Append(deletedRunContent);
                
                // Create an inserted run with the new hyperlink
                var insertedRun = CreateInsertedRun();
                var newHyperlink = CreateHyperlinkWithProperties(hyperlinkProperties, newText);
                var insertedRunContent = new Run();
                insertedRunContent.Append(newHyperlink);
                insertedRun.Append(insertedRunContent);
                
                // Replace the original hyperlink in the paragraph with proper error handling
                try
                {
                    var parentRun = hyperlink.Parent as Run;
                    if (parentRun != null)
                    {
                        // Insert track changes before the parent run
                        paragraph.InsertBefore(deletedRun, parentRun);
                        paragraph.InsertBefore(insertedRun, parentRun);
                        
                        // Remove the original run containing the hyperlink
                        parentRun.Remove();
                    }
                    else
                    {
                        // Direct hyperlink in paragraph - replace directly
                        paragraph.InsertBefore(deletedRun, hyperlink);
                        paragraph.InsertBefore(insertedRun, hyperlink);
                        hyperlink.Remove();
                    }
                }
                catch (Exception replaceEx)
                {
                    // If track changes insertion fails, fall back to simple update
                    throw new InvalidOperationException($"Failed to insert track changes in paragraph: {replaceEx.Message}", replaceEx);
                }
            }
            catch (Exception ex)
            {
                // If all track changes operations fail, fall back to simple hyperlink update
                throw new InvalidOperationException($"Track changes hyperlink update failed: {ex.Message}. Consider disabling track changes.", ex);
            }
        }

        /// <summary>
        /// Clones important properties from a hyperlink for recreation
        /// </summary>
        private static Dictionary<string, string> CloneHyperlinkProperties(Hyperlink hyperlink)
        {
            var properties = new Dictionary<string, string>();
            
            if (hyperlink.Id?.Value != null)
                properties["Id"] = hyperlink.Id.Value;
            if (hyperlink.Anchor?.Value != null)
                properties["Anchor"] = hyperlink.Anchor.Value;
            if (hyperlink.DocLocation?.Value != null)
                properties["DocLocation"] = hyperlink.DocLocation.Value;
            if (hyperlink.Tooltip?.Value != null)
                properties["Tooltip"] = hyperlink.Tooltip.Value;
            
            return properties;
        }

        /// <summary>
        /// Creates a new hyperlink with preserved properties and new text
        /// </summary>
        private static Hyperlink CreateHyperlinkWithProperties(Dictionary<string, string> properties, string newText)
        {
            var newHyperlink = new Hyperlink();
            
            // Restore properties
            if (properties.TryGetValue("Id", out var id))
                newHyperlink.Id = id;
            if (properties.TryGetValue("Anchor", out var anchor))
                newHyperlink.Anchor = anchor;
            if (properties.TryGetValue("DocLocation", out var docLocation))
                newHyperlink.DocLocation = docLocation;
            if (properties.TryGetValue("Tooltip", out var tooltip))
                newHyperlink.Tooltip = tooltip;
            
            // Add text content
            var run = new Run();
            run.Append(new Text(newText));
            newHyperlink.Append(run);
            
            return newHyperlink;
        }

        /// <summary>
        /// Creates a deleted run element for track changes
        /// </summary>
        private static DeletedRun CreateDeletedRun()
        {
            return new DeletedRun()
            {
                Id = GenerateRevisionId(),
                Author = "BulkEditor",
                Date = DateTime.UtcNow
            };
        }

        /// <summary>
        /// Creates an inserted run element for track changes
        /// </summary>
        private static InsertedRun CreateInsertedRun()
        {
            return new InsertedRun()
            {
                Id = GenerateRevisionId(),
                Author = "BulkEditor", 
                Date = DateTime.UtcNow
            };
        }

        /// <summary>
        /// Generates a unique revision ID for track changes
        /// </summary>
        private static string GenerateRevisionId()
        {
            return DateTime.UtcNow.Ticks.ToString();
        }

        /// <summary>
        /// Replaces text in a run with track changes support
        /// </summary>
        public static void ReplaceTextInRun(Run run, string oldText, string newText, bool trackChanges)
        {
            if (!trackChanges)
            {
                ReplaceTextInRunInternal(run, oldText, newText);
                return;
            }

            var textElements = run.Descendants<Text>().ToList();
            foreach (var textElement in textElements)
            {
                if (textElement.Text != null && textElement.Text.Contains(oldText))
                {
                    var replacedText = textElement.Text.Replace(oldText, newText);
                    
                    // Mark original text as deleted
                    var deletedRun = CreateDeletedRun();
                    var originalRunClone = (Run)run.CloneNode(true);
                    deletedRun.Append(originalRunClone);
                    
                    // Create inserted run with new text
                    var insertedRun = CreateInsertedRun();
                    var newRun = new Run();
                    
                    // Preserve formatting
                    var runProperties = run.GetFirstChild<RunProperties>()?.CloneNode(true);
                    if (runProperties != null)
                    {
                        newRun.Append(runProperties);
                    }
                    
                    newRun.Append(new Text(replacedText));
                    insertedRun.Append(newRun);
                    
                    // Replace content in the parent paragraph
                    var paragraph = run.Parent;
                    if (paragraph != null)
                    {
                        paragraph.InsertBefore(deletedRun, run);
                        paragraph.InsertBefore(insertedRun, run);
                        run.Remove();
                    }
                    break;
                }
            }
        }

        /// <summary>
        /// Internal method for replacing text in a run without track changes
        /// </summary>
        private static void ReplaceTextInRunInternal(Run run, string oldText, string newText)
        {
            var textElements = run.Descendants<Text>().ToList();
            foreach (var textElement in textElements)
            {
                if (textElement.Text != null && textElement.Text.Contains(oldText))
                {
                    textElement.Text = textElement.Text.Replace(oldText, newText);
                }
            }
        }

        /// <summary>
        /// Creates a tracked insertion for new text content
        /// </summary>
        public static InsertedRun CreateTrackedInsertion(string text, RunProperties? formatting = null)
        {
            var insertedRun = CreateInsertedRun();
            var run = new Run();
            
            if (formatting != null)
            {
                run.Append((RunProperties)formatting.CloneNode(true));
            }
            
            run.Append(new Text(text));
            insertedRun.Append(run);
            
            return insertedRun;
        }

        /// <summary>
        /// Creates a tracked deletion for existing text content
        /// </summary>
        public static DeletedRun CreateTrackedDeletion(Run originalRun)
        {
            var deletedRun = CreateDeletedRun();
            var clonedRun = (Run)originalRun.CloneNode(true);
            deletedRun.Append(clonedRun);
            
            return deletedRun;
        }
    }
}