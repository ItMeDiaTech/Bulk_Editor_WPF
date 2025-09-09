using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using BulkEditor.Core.Configuration;
using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using BulkEditor.UI.Services;
using BulkEditor.UI.Views;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;

namespace BulkEditor.UI.ViewModels
{
    /// <summary>
    /// Simplified ViewModel for the Processing Options window
    /// </summary>
    public partial class SimpleProcessingOptionsViewModel : ViewModelBase
    {
        private readonly ILoggingService _logger;
        private readonly INotificationService _notificationService;
        private readonly BulkEditor.Core.Services.IConfigurationService _configurationService;

        // Processing Options
        [ObservableProperty]
        private bool _updateTheSourceHyperlinkUrls = true;

        [ObservableProperty]
        private bool _appendContentIdsToTheSourceHyperlinks = true;

        [ObservableProperty]
        private bool _checkForExpiredContent = true;

        [ObservableProperty]
        private bool _autoReplaceOutdatedTitles = true;

        [ObservableProperty]
        private bool _reportTitleDifferencesInChangelog = true;

        [ObservableProperty]
        private bool _replaceCustomUserDefinedHyperlinks = false;

        [ObservableProperty]
        private bool _replaceCustomUserDefinedText = false;

        [ObservableProperty]
        private bool _optimizeTextFormatting = true;

        [ObservableProperty]
        private bool _createBackupBeforeProcessing = true;

        [ObservableProperty]
        private int _maxConcurrentDocuments = 5;

        [ObservableProperty]
        private int _timeoutPerDocumentMinutes = 10;

        // Replacement Rules
        [ObservableProperty]
        private ObservableCollection<HyperlinkReplacementRule> _hyperlinkRules = new();

        [ObservableProperty]
        private ObservableCollection<TextReplacementRule> _textRules = new();

        // Hyperlink matching mode for title comparison
        [ObservableProperty]
        private HyperlinkMatchMode _hyperlinkMatchMode = HyperlinkMatchMode.Contains;

        public SimpleProcessingOptionsViewModel(
            ILoggingService logger,
            INotificationService notificationService,
            BulkEditor.Core.Services.IConfigurationService configurationService)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _notificationService = notificationService ?? throw new ArgumentNullException(nameof(notificationService));
            _configurationService = configurationService ?? throw new ArgumentNullException(nameof(configurationService));

            InitializeDefaultSettings();
            // CRITICAL FIX: Don't load settings synchronously in constructor
            // Settings will be loaded asynchronously when the window is shown
        }

        private void InitializeDefaultSettings()
        {
            try
            {
                // Add some sample hyperlink rules
                HyperlinkRules.Add(new HyperlinkReplacementRule
                {
                    IsEnabled = true,
                    TitleToMatch = "User Guide",
                    ContentId = "CMS-DOC-123456",
                    CreatedAt = DateTime.Now
                });

                // Add some sample text rules
                TextRules.Add(new TextReplacementRule
                {
                    IsEnabled = true,
                    SourceText = "Old Company Name",
                    ReplacementText = "New Company Name",
                    CreatedAt = DateTime.Now
                });

                _logger.LogInformation("Processing options initialized with default settings");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error initializing processing options");
                _notificationService.ShowError("Initialization Error", "Failed to initialize processing options.", ex);
            }
        }

        /// <summary>
        /// CRITICAL FIX: Async method to load current settings without blocking UI thread
        /// </summary>
        public async Task LoadCurrentSettingsAsync()
        {
            try
            {
                _logger.LogDebug("Loading current settings asynchronously");
                var currentSettings = await _configurationService.LoadSettingsAsync().ConfigureAwait(false);

                // CRITICAL FIX: Ensure UI updates happen on UI thread
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    // Load processing settings into UI
                    UpdateTheSourceHyperlinkUrls = currentSettings.Processing.UpdateHyperlinks;
                    AppendContentIdsToTheSourceHyperlinks = currentSettings.Processing.AddContentIds;
                    CreateBackupBeforeProcessing = currentSettings.Processing.CreateBackupBeforeProcessing;
                    MaxConcurrentDocuments = currentSettings.Processing.MaxConcurrentDocuments;
                    TimeoutPerDocumentMinutes = (int)currentSettings.Processing.TimeoutPerDocument.TotalMinutes;
                    OptimizeTextFormatting = currentSettings.Processing.OptimizeText;

                    // Load validation settings
                    CheckForExpiredContent = currentSettings.Validation.CheckExpiredContent;
                    AutoReplaceOutdatedTitles = currentSettings.Validation.AutoReplaceTitles;
                    ReportTitleDifferencesInChangelog = currentSettings.Validation.ReportTitleDifferences;

                    // Load replacement settings
                    ReplaceCustomUserDefinedHyperlinks = currentSettings.Replacement.EnableHyperlinkReplacement;
                    ReplaceCustomUserDefinedText = currentSettings.Replacement.EnableTextReplacement;

                    // Load custom replacement rules
                    HyperlinkRules.Clear();
                    foreach (var rule in currentSettings.Replacement.HyperlinkRules)
                    {
                        HyperlinkRules.Add(rule);
                    }

                    TextRules.Clear();
                    foreach (var rule in currentSettings.Replacement.TextRules)
                    {
                        TextRules.Add(rule);
                    }

                    // Set hyperlink match mode (default to Contains for backward compatibility)
                    HyperlinkMatchMode = HyperlinkMatchMode.Contains;
                });

                _logger.LogDebug("Current settings loaded successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error loading current settings asynchronously");
                _notificationService.ShowError("Settings Load Error", "Failed to load current settings. Using defaults.", ex);
            }
        }

        /// <summary>
        /// CRITICAL FIX: Async method to save settings without blocking UI thread
        /// </summary>
        [RelayCommand]
        private async Task SaveSettingsAsync()
        {
            try
            {
                _logger.LogDebug("Saving processing options asynchronously");

                // Get current app settings asynchronously
                var currentSettings = await _configurationService.LoadSettingsAsync().ConfigureAwait(false);

                // Update processing settings from UI values
                currentSettings.Processing.UpdateHyperlinks = UpdateTheSourceHyperlinkUrls;
                currentSettings.Processing.AddContentIds = AppendContentIdsToTheSourceHyperlinks;
                currentSettings.Processing.CreateBackupBeforeProcessing = CreateBackupBeforeProcessing;
                currentSettings.Processing.MaxConcurrentDocuments = MaxConcurrentDocuments;
                currentSettings.Processing.TimeoutPerDocument = TimeSpan.FromMinutes(TimeoutPerDocumentMinutes);
                currentSettings.Processing.OptimizeText = OptimizeTextFormatting;

                // Update validation settings
                currentSettings.Validation.CheckExpiredContent = CheckForExpiredContent;
                currentSettings.Validation.AutoReplaceTitles = AutoReplaceOutdatedTitles;
                currentSettings.Validation.ReportTitleDifferences = ReportTitleDifferencesInChangelog;

                // Update replacement settings
                currentSettings.Replacement.EnableHyperlinkReplacement = ReplaceCustomUserDefinedHyperlinks;
                currentSettings.Replacement.EnableTextReplacement = ReplaceCustomUserDefinedText;

                // Save custom replacement rules
                currentSettings.Replacement.HyperlinkRules = HyperlinkRules.ToList();
                currentSettings.Replacement.TextRules = TextRules.ToList();

                // Save the updated settings asynchronously
                await _configurationService.SaveSettingsAsync(currentSettings).ConfigureAwait(false);

                _logger.LogInformation("Processing options saved successfully");

                // CRITICAL FIX: UI updates must happen on UI thread
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    _notificationService.ShowSuccess("Settings Saved", "Processing options have been saved successfully.");

                    // Close the window with OK result
                    if (System.Windows.Application.Current.Windows.OfType<ProcessingOptionsWindow>().FirstOrDefault() is ProcessingOptionsWindow window)
                    {
                        window.DialogResult = true;
                        window.Close();
                    }
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving processing options asynchronously");
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    _notificationService.ShowError("Save Error", "Failed to save processing options.", ex);
                });
            }
        }

        [RelayCommand]
        private void Cancel()
        {
            if (System.Windows.Application.Current.Windows.OfType<ProcessingOptionsWindow>().FirstOrDefault() is ProcessingOptionsWindow window)
            {
                window.DialogResult = false;
                window.Close();
            }
        }

        [RelayCommand]
        private void ResetToDefaults()
        {
            var result = MessageBox.Show(
                "Are you sure you want to reset all processing options to their default values? This action cannot be undone.",
                "Reset to Defaults",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                // Reset processing options to defaults
                UpdateTheSourceHyperlinkUrls = true;
                AppendContentIdsToTheSourceHyperlinks = true;
                CheckForExpiredContent = true;
                AutoReplaceOutdatedTitles = true;
                ReportTitleDifferencesInChangelog = true;
                ReplaceCustomUserDefinedHyperlinks = false;
                ReplaceCustomUserDefinedText = false;
                OptimizeTextFormatting = true;
                CreateBackupBeforeProcessing = true;
                MaxConcurrentDocuments = 5;
                TimeoutPerDocumentMinutes = 10;

                // Clear all rules
                HyperlinkRules.Clear();
                TextRules.Clear();

                _logger.LogInformation("Processing options reset to defaults");
                _notificationService.ShowInfo("Reset Complete", "Processing options have been reset to default values.");
            }
        }

        #region Hyperlink Rules Commands

        [RelayCommand]
        private void AddHyperlinkRule()
        {
            var newRule = new HyperlinkReplacementRule
            {
                IsEnabled = true,
                TitleToMatch = "New Rule",
                ContentId = "",
                CreatedAt = DateTime.Now
            };

            HyperlinkRules.Add(newRule);
            _logger.LogInformation("New hyperlink rule added");
        }

        [RelayCommand]
        private void RemoveHyperlinkRule(HyperlinkReplacementRule? rule)
        {
            if (rule != null && HyperlinkRules.Contains(rule))
            {
                HyperlinkRules.Remove(rule);
                _logger.LogInformation("Hyperlink rule removed: {Rule}", rule.TitleToMatch);
            }
        }

        [RelayCommand]
        private async Task ExportHyperlinkRulesAsync()
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Title = "Export Hyperlink Rules",
                    Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                    DefaultExt = "json",
                    FileName = $"HyperlinkRules_{DateTime.Now:yyyyMMdd_HHmmss}.json"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    await System.IO.File.WriteAllTextAsync(saveFileDialog.FileName,
                        System.Text.Json.JsonSerializer.Serialize(HyperlinkRules.ToList(), new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));

                    _notificationService.ShowSuccess("Export Complete", $"Hyperlink rules exported to {System.IO.Path.GetFileName(saveFileDialog.FileName)}");
                    _logger.LogInformation("Hyperlink rules exported to {FilePath}", saveFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting hyperlink rules");
                _notificationService.ShowError("Export Error", "Failed to export hyperlink rules.", ex);
            }
        }

        [RelayCommand]
        private async Task ImportHyperlinkRulesAsync()
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Title = "Import Hyperlink Rules",
                    Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                    CheckFileExists = true
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    var jsonContent = await System.IO.File.ReadAllTextAsync(openFileDialog.FileName);
                    var importedRules = System.Text.Json.JsonSerializer.Deserialize<HyperlinkReplacementRule[]>(jsonContent);

                    if (importedRules != null)
                    {
                        foreach (var rule in importedRules)
                        {
                            HyperlinkRules.Add(rule);
                        }

                        _notificationService.ShowSuccess("Import Complete", $"Imported {importedRules.Length} hyperlink rules");
                        _logger.LogInformation("Imported {Count} hyperlink rules from {FilePath}", importedRules.Length, openFileDialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error importing hyperlink rules");
                _notificationService.ShowError("Import Error", "Failed to import hyperlink rules.", ex);
            }
        }

        [RelayCommand]
        private void ClearAllHyperlinkRules()
        {
            var result = MessageBox.Show(
                "Are you sure you want to clear all hyperlink rules? This action cannot be undone.",
                "Clear All Rules",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                var count = HyperlinkRules.Count;
                HyperlinkRules.Clear();
                _notificationService.ShowInfo("Rules Cleared", $"Cleared {count} hyperlink rules");
                _logger.LogInformation("Cleared {Count} hyperlink rules", count);
            }
        }

        #endregion

        #region Text Rules Commands

        [RelayCommand]
        private void AddTextRule()
        {
            var newRule = new TextReplacementRule
            {
                IsEnabled = true,
                SourceText = "Source Text",
                ReplacementText = "Replacement Text",
                CreatedAt = DateTime.Now
            };

            TextRules.Add(newRule);
            _logger.LogInformation("New text rule added");
        }

        [RelayCommand]
        private void RemoveTextRule(TextReplacementRule? rule)
        {
            if (rule != null && TextRules.Contains(rule))
            {
                TextRules.Remove(rule);
                _logger.LogInformation("Text rule removed: {Rule}", rule.SourceText);
            }
        }

        [RelayCommand]
        private async Task ExportTextRulesAsync()
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Title = "Export Text Rules",
                    Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                    DefaultExt = "json",
                    FileName = $"TextRules_{DateTime.Now:yyyyMMdd_HHmmss}.json"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    await System.IO.File.WriteAllTextAsync(saveFileDialog.FileName,
                        System.Text.Json.JsonSerializer.Serialize(TextRules.ToList(), new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));

                    _notificationService.ShowSuccess("Export Complete", $"Text rules exported to {System.IO.Path.GetFileName(saveFileDialog.FileName)}");
                    _logger.LogInformation("Text rules exported to {FilePath}", saveFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting text rules");
                _notificationService.ShowError("Export Error", "Failed to export text rules.", ex);
            }
        }

        [RelayCommand]
        private async Task ImportTextRulesAsync()
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Title = "Import Text Rules",
                    Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                    CheckFileExists = true
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    var jsonContent = await System.IO.File.ReadAllTextAsync(openFileDialog.FileName);
                    var importedRules = System.Text.Json.JsonSerializer.Deserialize<TextReplacementRule[]>(jsonContent);

                    if (importedRules != null)
                    {
                        foreach (var rule in importedRules)
                        {
                            TextRules.Add(rule);
                        }

                        _notificationService.ShowSuccess("Import Complete", $"Imported {importedRules.Length} text rules");
                        _logger.LogInformation("Imported {Count} text rules from {FilePath}", importedRules.Length, openFileDialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error importing text rules");
                _notificationService.ShowError("Import Error", "Failed to import text rules.", ex);
            }
        }

        [RelayCommand]
        private void ClearAllTextRules()
        {
            var result = MessageBox.Show(
                "Are you sure you want to clear all text rules? This action cannot be undone.",
                "Clear All Rules",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                var count = TextRules.Count;
                TextRules.Clear();
                _notificationService.ShowInfo("Rules Cleared", $"Cleared {count} text rules");
                _logger.LogInformation("Cleared {Count} text rules", count);
            }
        }

        #endregion

        public override void Cleanup()
        {
            base.Cleanup();
        }
    }

    /// <summary>
    /// Enum for hyperlink title matching modes
    /// </summary>
    public enum HyperlinkMatchMode
    {
        /// <summary>Exact title match (case-insensitive)</summary>
        Exact,
        /// <summary>Title contains the match text (case-insensitive)</summary>
        Contains,
        /// <summary>Title starts with the match text (case-insensitive)</summary>
        StartsWith,
        /// <summary>Title ends with the match text (case-insensitive)</summary>
        EndsWith
    }
}