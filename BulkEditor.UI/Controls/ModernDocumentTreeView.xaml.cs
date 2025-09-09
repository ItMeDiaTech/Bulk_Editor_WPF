using System.Windows;
using System.Windows.Controls;
using BulkEditor.UI.ViewModels;

namespace BulkEditor.UI.Controls
{
    /// <summary>
    /// Modern expandable tree view control for displaying document processing results
    /// </summary>
    public partial class ModernDocumentTreeView : UserControl
    {
        public ModernDocumentTreeView()
        {
            InitializeComponent();
        }

        private void ExpandAllButton_Click(object sender, RoutedEventArgs e)
        {
            SetTreeViewItemsExpanded(true);
        }

        private void CollapseAllButton_Click(object sender, RoutedEventArgs e)
        {
            SetTreeViewItemsExpanded(false);
        }

        private void SetTreeViewItemsExpanded(bool isExpanded)
        {
            if (DataContext is MainWindowViewModel viewModel)
            {
                foreach (var result in viewModel.ProcessingResults)
                {
                    result.IsExpanded = isExpanded;
                    
                    foreach (var option in result.ProcessingOptions)
                    {
                        option.IsExpanded = isExpanded;
                    }
                }
            }
        }

        /// <summary>
        /// Expands tree view items to show a specific level of detail
        /// </summary>
        /// <param name="maxLevel">Maximum level to expand (1 = documents only, 2 = documents + options, 3 = all)</param>
        public void ExpandToLevel(int maxLevel)
        {
            if (DataContext is not MainWindowViewModel viewModel) return;

            foreach (var result in viewModel.ProcessingResults)
            {
                // Level 1: Always expand documents
                result.IsExpanded = maxLevel >= 1;
                
                if (maxLevel >= 2)
                {
                    // Level 2: Expand processing options that have changes
                    foreach (var option in result.ProcessingOptions)
                    {
                        option.IsExpanded = option.HasChanges && maxLevel >= 3;
                    }
                }
            }
        }
    }
}