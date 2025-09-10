using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Linq;

namespace BulkEditor.Infrastructure.Utilities
{
    public static class OpenXmlHelper
    {
        /// <summary>
        /// Safely updates hyperlink text, preserving formatting and hyperlink properties (DocLocation, Anchor, etc.).
        /// </summary>
        public static void UpdateHyperlinkText(Hyperlink hyperlink, string newText)
        {
            // Preserve existing formatting by finding the first run and cloning its properties
            var runProperties = hyperlink.Descendants<Run>()
                .FirstOrDefault()
                ?.GetFirstChild<RunProperties>()
                ?.CloneNode(true);

            // CRITICAL FIX: Preserve important hyperlink properties before removing children
            var docLocation = hyperlink.DocLocation?.Value;
            var anchor = hyperlink.Anchor?.Value;
            var tooltip = hyperlink.Tooltip?.Value;
            var history = hyperlink.History?.Value;

            // Clear existing content and add the new, properly structured text
            hyperlink.RemoveAllChildren();
            
            // Re-add preserved properties first
            if (docLocation != null)
                hyperlink.DocLocation = new StringValue(docLocation);
            if (anchor != null)
                hyperlink.Anchor = new StringValue(anchor);
            if (tooltip != null)
                hyperlink.Tooltip = new StringValue(tooltip);
            if (history != null)
                hyperlink.History = new OnOffValue(history);
            
            // Add the new text run
            var newRun = new Run();
            if (runProperties != null)
            {
                newRun.Append(runProperties);
            }
            newRun.Append(new Text(newText));
            hyperlink.Append(newRun);
        }
    }
}