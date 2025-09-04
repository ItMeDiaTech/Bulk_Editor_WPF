using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using BulkEditor.Core.Configuration;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;

namespace BulkEditor.UI.ViewModels.Settings
{
    public partial class ReplacementSettingsViewModel : ObservableObject
    {
        [ObservableProperty]
        private bool _enableHyperlinkReplacement;

        [ObservableProperty]
        private bool _enableTextReplacement;

        [ObservableProperty]
        private int _maxReplacementRules;

        [ObservableProperty]
        private bool _validateContentIds;

        public ObservableCollection<HyperlinkReplacementRule> HyperlinkRules { get; set; } = new();
        public ObservableCollection<TextReplacementRule> TextRules { get; set; } = new();

        [RelayCommand]
        private void AddHyperlinkRule()
        {
            var newRule = new HyperlinkReplacementRule
            {
                Id = Guid.NewGuid().ToString(),
                TitleToMatch = "Enter title to match",
                ContentId = "Enter content ID",
                IsEnabled = true,
                CreatedAt = DateTime.UtcNow
            };
            HyperlinkRules.Add(newRule);
        }

        [RelayCommand]
        private void RemoveHyperlinkRule(HyperlinkReplacementRule? rule)
        {
            if (rule != null && HyperlinkRules.Contains(rule))
            {
                HyperlinkRules.Remove(rule);
            }
        }

        [RelayCommand]
        private void AddTextRule()
        {
            var newRule = new TextReplacementRule
            {
                Id = Guid.NewGuid().ToString(),
                SourceText = "Enter source text",
                ReplacementText = "Enter replacement text",
                IsEnabled = true,
                CreatedAt = DateTime.UtcNow
            };
            TextRules.Add(newRule);
        }

        [RelayCommand]
        private void RemoveTextRule(TextReplacementRule? rule)
        {
            if (rule != null && TextRules.Contains(rule))
            {
                TextRules.Remove(rule);
            }
        }

        [RelayCommand]
        private void ExportRules()
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Title = "Export Replacement Rules",
                    Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                    DefaultExt = "json",
                    FileName = $"ReplacementRules_{DateTime.Now:yyyyMMdd_HHmmss}.json"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    var exportData = new
                    {
                        HyperlinkRules = HyperlinkRules.ToList(),
                        TextRules = TextRules.ToList(),
                        ExportedAt = DateTime.UtcNow,
                        Version = "1.0"
                    };

                    var json = JsonSerializer.Serialize(exportData, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(saveDialog.FileName, json);

                    System.Windows.MessageBox.Show(
                        $"Rules exported successfully to:\n{saveDialog.FileName}",
                        "Export Successful",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"Failed to export rules:\n{ex.Message}",
                    "Export Failed",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
        }

        [RelayCommand]
        private void ImportRules()
        {
            try
            {
                var openDialog = new OpenFileDialog
                {
                    Title = "Import Replacement Rules",
                    Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                    DefaultExt = "json"
                };

                if (openDialog.ShowDialog() == true)
                {
                    var json = File.ReadAllText(openDialog.FileName);
                    var importData = JsonSerializer.Deserialize<JsonElement>(json);

                    if (importData.TryGetProperty("HyperlinkRules", out var hyperlinkRulesElement))
                    {
                        var hyperlinkRules = JsonSerializer.Deserialize<HyperlinkReplacementRule[]>(hyperlinkRulesElement.GetRawText());
                        if (hyperlinkRules != null)
                        {
                            foreach (var rule in hyperlinkRules)
                            {
                                HyperlinkRules.Add(rule);
                            }
                        }
                    }

                    if (importData.TryGetProperty("TextRules", out var textRulesElement))
                    {
                        var textRules = JsonSerializer.Deserialize<TextReplacementRule[]>(textRulesElement.GetRawText());
                        if (textRules != null)
                        {
                            foreach (var rule in textRules)
                            {
                                TextRules.Add(rule);
                            }
                        }
                    }

                    System.Windows.MessageBox.Show(
                        "Rules imported successfully!",
                        "Import Successful",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"Failed to import rules:\n{ex.Message}",
                    "Import Failed",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
        }

        [RelayCommand]
        private void ClearAllRules()
        {
            var result = System.Windows.MessageBox.Show(
                "Are you sure you want to clear all replacement rules?\nThis action cannot be undone.",
                "Clear All Rules",
                System.Windows.MessageBoxButton.YesNo,
                System.Windows.MessageBoxImage.Warning);

            if (result == System.Windows.MessageBoxResult.Yes)
            {
                HyperlinkRules.Clear();
                TextRules.Clear();

                System.Windows.MessageBox.Show(
                    "All replacement rules have been cleared.",
                    "Rules Cleared",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
            }
        }
    }
}