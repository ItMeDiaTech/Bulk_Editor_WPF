using System.Windows;
using System.Windows.Controls;
using BulkEditor.UI.ViewModels;

namespace BulkEditor.UI.Selectors
{
    public class DocumentViewTemplateSelector : DataTemplateSelector
    {
        public DataTemplate? DetailedTemplate { get; set; }
        public DataTemplate? CompactTemplate { get; set; }

        public override DataTemplate? SelectTemplate(object item, DependencyObject container)
        {
            if (item is DocumentListItemViewModel && container != null)
            {
                // Get the MainWindowViewModel from the window
                var window = Window.GetWindow(container);
                if (window?.DataContext is MainWindowViewModel viewModel)
                {
                    return viewModel.IsCompactView ? CompactTemplate : DetailedTemplate;
                }
            }

            return DetailedTemplate ?? base.SelectTemplate(item, container);
        }
    }
}