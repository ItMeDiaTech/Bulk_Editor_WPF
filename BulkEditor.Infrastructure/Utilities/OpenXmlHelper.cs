using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace BulkEditor.Infrastructure.Utilities
{
    public static class OpenXmlHelper
    {
        /// <summary>
        /// Safely updates hyperlink text, preserving formatting by wrapping text in a Run.
        /// </summary>
        public static void UpdateHyperlinkText(Hyperlink hyperlink, string newText)
        {
            // Preserve existing formatting by finding the first run and cloning its properties
            var runProperties = hyperlink.Descendants<Run>()
                .FirstOrDefault()
                ?.GetFirstChild<RunProperties>()
                ?.CloneNode(true);

            // Clear existing content and add the new, properly structured text
            hyperlink.RemoveAllChildren();
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