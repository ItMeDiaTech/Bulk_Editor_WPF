using System;
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

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // CRITICAL FIX: Load settings asynchronously when window loads
            if (DataContext is SimpleProcessingOptionsViewModel viewModel)
            {
                try
                {
                    // Load current settings asynchronously to prevent UI freeze
                    await viewModel.LoadCurrentSettingsAsync();
                }
                catch (Exception ex)
                {
                    // Log error but don't crash the window
                    System.Diagnostics.Debug.WriteLine($"Error loading settings: {ex.Message}");
                }
            }
        }
    }
}