using BulkEditor.UI.ViewModels;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace BulkEditor.UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly MainWindowViewModel _viewModel;
        private Brush? _originalBorderBrush;
        private Brush? _originalBackground;

        public MainWindow(MainWindowViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = viewModel;

            // Initialize the ViewModel
            Loaded += async (s, e) => await viewModel.InitializeAsync();

            // Cleanup when closing
            Closing += (s, e) => viewModel.Cleanup();
        }

        #region Drag & Drop Event Handlers

        private void DocumentsArea_DragEnter(object sender, DragEventArgs e)
        {
            // Store original appearance
            if (sender is Border border)
            {
                _originalBorderBrush = border.BorderBrush;
                _originalBackground = border.Background;
            }

            if (HasValidFiles(e))
            {
                e.Effects = DragDropEffects.Copy;

                // Visual feedback for valid drop
                if (sender is Border validBorder)
                {
                    validBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(0x21, 0x96, 0xF3));
                    validBorder.Background = new SolidColorBrush(Color.FromArgb(20, 0x21, 0x96, 0xF3));
                }
            }
            else
            {
                e.Effects = DragDropEffects.None;

                // Visual feedback for invalid drop
                if (sender is Border invalidBorder)
                {
                    invalidBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(0xF4, 0x43, 0x36));
                    invalidBorder.Background = new SolidColorBrush(Color.FromArgb(20, 0xF4, 0x43, 0x36));
                }
            }

            e.Handled = true;
        }

        private void DocumentsArea_DragOver(object sender, DragEventArgs e)
        {
            if (HasValidFiles(e))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }

            e.Handled = true;
        }

        private void DocumentsArea_DragLeave(object sender, DragEventArgs e)
        {
            // Restore original appearance
            if (sender is Border border)
            {
                border.BorderBrush = _originalBorderBrush;
                border.Background = _originalBackground;
            }

            e.Handled = true;
        }

        private async void DocumentsArea_Drop(object sender, DragEventArgs e)
        {
            // Restore original appearance
            if (sender is Border border)
            {
                border.BorderBrush = _originalBorderBrush;
                border.Background = _originalBackground;
            }

            if (!HasValidFiles(e))
            {
                e.Handled = true;
                return;
            }

            try
            {
                var files = GetDroppedFiles(e);
                if (files?.Any() == true)
                {
                    await _viewModel.AddFilesCommand.ExecuteAsync(files);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error processing dropped files: {ex.Message}", "Drag & Drop Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }

            e.Handled = true;
        }

        #endregion

        #region Helper Methods

        private bool HasValidFiles(DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return false;

            var files = GetDroppedFiles(e);
            return files?.Any(IsValidWordDocument) == true;
        }

        private string[]? GetDroppedFiles(DragEventArgs e)
        {
            return e.Data.GetData(DataFormats.FileDrop) as string[];
        }

        private bool IsValidWordDocument(string filePath)
        {
            if (!File.Exists(filePath))
                return false;

            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension == ".docx" || extension == ".docm";
        }

        #endregion
    }
}