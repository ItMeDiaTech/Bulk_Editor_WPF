using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Linq;

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
    }
}