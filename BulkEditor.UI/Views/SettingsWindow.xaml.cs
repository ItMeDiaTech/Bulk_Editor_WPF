using BulkEditor.UI.ViewModels;
using System;
using System.Windows;

namespace BulkEditor.UI.Views
{
    /// <summary>
    /// Interaction logic for SettingsWindow.xaml
    /// </summary>
    public partial class SettingsWindow : Window
    {
        private readonly SettingsViewModel _viewModel;

        public SettingsWindow(SettingsViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = viewModel;

            // Handle dialog result from ViewModel
            _viewModel.RequestClose += ViewModel_RequestClose;
        }

        private void ViewModel_RequestClose(object? sender, bool? result)
        {
            DialogResult = result;
            Close();
        }

        protected override void OnClosed(EventArgs e)
        {
            // Unsubscribe from event to prevent memory leaks
            _viewModel.RequestClose -= ViewModel_RequestClose;
            base.OnClosed(e);
        }
    }
}