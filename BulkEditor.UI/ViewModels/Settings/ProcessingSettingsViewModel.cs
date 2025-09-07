using BulkEditor.Core.Interfaces;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;

namespace BulkEditor.UI.ViewModels.Settings
{
    public partial class ProcessingSettingsViewModel : ObservableObject
    {
        [ObservableProperty]
        private int _maxConcurrentDocuments;

        [ObservableProperty]
        private int _batchSize;

        [ObservableProperty]
        private bool _createBackupBeforeProcessing;


        [ObservableProperty]
        private int _timeoutPerDocumentMinutes;

        [ObservableProperty]
        private string _consultantEmail = string.Empty;

        [ObservableProperty]
        private string _selectedTheme = "Light";

        private readonly IThemeService? _themeService;

        public ObservableCollection<string> AvailableThemes { get; } = new();

        public ProcessingSettingsViewModel(IThemeService? themeService = null)
        {
            _themeService = themeService;
            
            // Initialize available themes
            if (_themeService != null)
            {
                foreach (var theme in _themeService.AvailableThemes)
                {
                    AvailableThemes.Add(theme);
                }
                
                // Set current theme
                SelectedTheme = _themeService.CurrentTheme;
            }
            else
            {
                // Fallback if theme service is not available
                AvailableThemes.Add("Light");
                AvailableThemes.Add("Dark");
                AvailableThemes.Add("Auto");
            }
        }

        partial void OnSelectedThemeChanged(string value)
        {
            if (_themeService != null && !string.IsNullOrEmpty(value))
            {
                _ = Task.Run(async () => await _themeService.ApplyThemeAsync(value));
            }
        }

        [RelayCommand]
        private async Task ExportAllSettings()
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Title = "Export Settings",
                    Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                    DefaultExt = "json",
                    FileName = $"BulkEditor_Settings_{DateTime.Now:yyyy-MM-dd_HHmm}.json"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    // TODO: Collect all settings from different ViewModels
                    var allSettings = new
                    {
                        Processing = new
                        {
                            MaxConcurrentDocuments,
                            BatchSize,
                            CreateBackupBeforeProcessing,
                            TimeoutPerDocumentMinutes,
                            ConsultantEmail
                        },
                        ExportedAt = DateTime.UtcNow,
                        Version = "1.0"
                    };

                    var json = JsonSerializer.Serialize(allSettings, new JsonSerializerOptions { WriteIndented = true });
                    await File.WriteAllTextAsync(saveDialog.FileName, json);
                    
                    MessageBox.Show("Settings exported successfully!", "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting settings: {ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        [RelayCommand]
        private async Task ImportAllSettings()
        {
            try
            {
                var openDialog = new OpenFileDialog
                {
                    Title = "Import Settings",
                    Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                    DefaultExt = "json"
                };

                if (openDialog.ShowDialog() == true)
                {
                    var json = await File.ReadAllTextAsync(openDialog.FileName);
                    using var document = JsonDocument.Parse(json);
                    
                    if (document.RootElement.TryGetProperty("Processing", out var processingElement))
                    {
                        if (processingElement.TryGetProperty("MaxConcurrentDocuments", out var maxConcurrent))
                            MaxConcurrentDocuments = maxConcurrent.GetInt32();
                        
                        if (processingElement.TryGetProperty("BatchSize", out var batchSize))
                            BatchSize = batchSize.GetInt32();
                        
                        if (processingElement.TryGetProperty("CreateBackupBeforeProcessing", out var createBackup))
                            CreateBackupBeforeProcessing = createBackup.GetBoolean();
                        
                        if (processingElement.TryGetProperty("TimeoutPerDocumentMinutes", out var timeout))
                            TimeoutPerDocumentMinutes = timeout.GetInt32();
                        
                        if (processingElement.TryGetProperty("ConsultantEmail", out var email))
                            ConsultantEmail = email.GetString() ?? string.Empty;
                    }
                    
                    MessageBox.Show("Settings imported successfully!", "Import Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error importing settings: {ex.Message}", "Import Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        [RelayCommand]
        private void ResetAllSettings()
        {
            var result = MessageBox.Show("Are you sure you want to reset all settings to defaults? This action cannot be undone.", 
                                       "Reset Settings", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            
            if (result == MessageBoxResult.Yes)
            {
                MaxConcurrentDocuments = 5;
                BatchSize = 100;
                CreateBackupBeforeProcessing = true;
                TimeoutPerDocumentMinutes = 5;
                ConsultantEmail = string.Empty;
                
                MessageBox.Show("Settings have been reset to defaults.", "Reset Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
}