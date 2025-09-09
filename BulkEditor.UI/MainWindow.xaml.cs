using BulkEditor.Core.Configuration;
using BulkEditor.UI.ViewModels;
using Microsoft.Extensions.DependencyInjection;
using System;
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

            // Load window position
            LoadWindowSettings();

            // Cleanup and save window position when closing
            Closing += (s, e) => 
            {
                SaveWindowSettings();
                viewModel.Cleanup();
            };
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

        #region Window Position Management

        private void LoadWindowSettings()
        {
            try
            {
                // Get AppSettings from DI container through Application
                var app = System.Windows.Application.Current as App;
                var serviceProvider = app?.ServiceProvider;
                var appSettings = serviceProvider?.GetService<AppSettings>();
                
                if (appSettings?.UI?.Window != null)
                {
                    var windowSettings = appSettings.UI.Window;
                    
                    // Set window size
                    if (windowSettings.Width > 0 && windowSettings.Height > 0)
                    {
                        Width = windowSettings.Width;
                        Height = windowSettings.Height;
                    }
                    
                    // Set window position (ensure it's on screen)
                    if (IsPositionOnScreen(windowSettings.Left, windowSettings.Top, windowSettings.Width, windowSettings.Height))
                    {
                        Left = windowSettings.Left;
                        Top = windowSettings.Top;
                    }
                    else
                    {
                        // Center window if saved position is off-screen
                        WindowStartupLocation = WindowStartupLocation.CenterScreen;
                    }
                    
                    // Set window state
                    if (windowSettings.Maximized)
                    {
                        WindowState = WindowState.Maximized;
                    }
                }
            }
            catch (Exception ex)
            {
                // If loading fails, use default centering
                WindowStartupLocation = WindowStartupLocation.CenterScreen;
                System.Diagnostics.Debug.WriteLine($"Failed to load window settings: {ex.Message}");
            }
        }
        
        private void SaveWindowSettings()
        {
            try
            {
                // Get AppSettings from DI container through Application
                var app = System.Windows.Application.Current as App;
                var serviceProvider = app?.ServiceProvider;
                var appSettings = serviceProvider?.GetService<AppSettings>();
                
                if (appSettings?.UI?.Window != null)
                {
                    var windowSettings = appSettings.UI.Window;
                    
                    // Save maximized state
                    windowSettings.Maximized = WindowState == WindowState.Maximized;
                    
                    // Save position and size (but only if not maximized)
                    if (WindowState == WindowState.Normal)
                    {
                        windowSettings.Width = ActualWidth;
                        windowSettings.Height = ActualHeight;
                        windowSettings.Left = Left;
                        windowSettings.Top = Top;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to save window settings: {ex.Message}");
            }
        }
        
        private bool IsPositionOnScreen(double left, double top, double width, double height)
        {
            try
            {
                // Simple bounds check against primary screen working area
                var workingArea = SystemParameters.WorkArea;
                
                // Check if at least part of the window would be visible
                return left < workingArea.Right - 100 && 
                       top < workingArea.Bottom - 100 && 
                       left + width > workingArea.Left + 100 && 
                       top + height > workingArea.Top + 100;
            }
            catch
            {
                // If we can't determine, assume it's safe
                return true;
            }
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