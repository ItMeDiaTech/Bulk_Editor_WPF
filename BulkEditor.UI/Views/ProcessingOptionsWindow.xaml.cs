using System.Windows;
using BulkEditor.UI.ViewModels;

namespace BulkEditor.UI.Views
{
    /// <summary>
    /// Interaction logic for ProcessingOptionsWindow.xaml
    /// </summary>
    public partial class ProcessingOptionsWindow : Window
    {
        public ProcessingOptionsWindow()
        {
            InitializeComponent();
        }

        public ProcessingOptionsWindow(SimpleProcessingOptionsViewModel viewModel) : this()
        {
            DataContext = viewModel;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Set focus to the first tab
            if (DataContext is SimpleProcessingOptionsViewModel viewModel)
            {
                // Any initialization needed when window loads
            }
        }
    }
}