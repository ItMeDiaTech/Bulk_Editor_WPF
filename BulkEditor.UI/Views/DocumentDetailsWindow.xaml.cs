using System.Windows;
using BulkEditor.UI.ViewModels;

namespace BulkEditor.UI.Views
{
    /// <summary>
    /// Interaction logic for DocumentDetailsWindow.xaml
    /// </summary>
    public partial class DocumentDetailsWindow : Window
    {
        public DocumentDetailsWindow()
        {
            InitializeComponent();
        }

        public DocumentDetailsWindow(DocumentDetailsViewModel viewModel) : this()
        {
            DataContext = viewModel;
            
            // Subscribe to close request from view model
            if (viewModel != null)
            {
                viewModel.CloseRequested += (s, e) => Close();
            }
        }
    }
}