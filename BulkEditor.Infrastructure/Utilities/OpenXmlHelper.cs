using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Linq;
using System;

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

            // Get original text for track changes
            var originalText = GetHyperlinkText(hyperlink);
            
            // Mark original content as deleted if it exists
            if (!string.IsNullOrEmpty(originalText))
            {
                MarkHyperlinkContentAsDeleted(hyperlink);
            }
            
            // Add new content as inserted
            AddInsertedTextToHyperlink(hyperlink, newText);
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
        /// Marks existing hyperlink content as deleted for track changes
        /// </summary>
        private static void MarkHyperlinkContentAsDeleted(Hyperlink hyperlink)
        {
            var existingRuns = hyperlink.Elements<Run>().ToList();
            
            if (existingRuns.Any())
            {
                // Create a deleted run containing all existing content
                var deletedRun = CreateDeletedRun();
                
                // Move existing runs into the deleted run
                foreach (var run in existingRuns)
                {
                    var clonedRun = (Run)run.CloneNode(true);
                    deletedRun.Append(clonedRun);
                }
                
                // Remove original runs
                foreach (var run in existingRuns)
                {
                    run.Remove();
                }
                
                // Add the deleted run to hyperlink
                hyperlink.Append(deletedRun);
            }
        }

        /// <summary>
        /// Adds new text as inserted content to hyperlink
        /// </summary>
        private static void AddInsertedTextToHyperlink(Hyperlink hyperlink, string newText)
        {
            var insertedRun = CreateInsertedRun();
            
            // Preserve formatting from original content if available
            var runProperties = hyperlink.Descendants<Run>()
                .FirstOrDefault()
                ?.GetFirstChild<RunProperties>()
                ?.CloneNode(true);
            
            var newRun = new Run();
            if (runProperties != null)
            {
                newRun.Append(runProperties);
            }
            newRun.Append(new Text(newText));
            
            insertedRun.Append(newRun);
            hyperlink.Append(insertedRun);
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