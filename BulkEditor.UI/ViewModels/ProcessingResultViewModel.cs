using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using BulkEditor.Core.Entities;

namespace BulkEditor.UI.ViewModels
{
    /// <summary>
    /// ViewModel representing a processing result for a single document with expandable sections
    /// </summary>
    public class ProcessingResultViewModel : INotifyPropertyChanged
    {
        private bool _isExpanded;
        private Document _document;

        public ProcessingResultViewModel(Document document)
        {
            _document = document;
            InitializeProcessingOptions();
        }

        public string DocumentTitle => _document.FileName;
        public string DocumentPath => _document.FilePath;
        public DocumentStatus Status => _document.Status;
        public DateTime? ProcessedAt => _document.ProcessedAt;
        
        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                _isExpanded = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<ProcessingOptionViewModel> ProcessingOptions { get; } = new();

        private void InitializeProcessingOptions()
        {
            // Group changes by processing option type
            var urlUpdates = _document.ChangeLog.Changes
                .Where(c => c.Type == Core.Entities.ChangeType.HyperlinkUpdated)
                .ToList();
            
            var contentIdAdditions = _document.ChangeLog.Changes
                .Where(c => c.Type == Core.Entities.ChangeType.ContentIdAdded)
                .ToList();
            
            var expiredContent = _document.ChangeLog.Changes
                .Where(c => c.Type == Core.Entities.ChangeType.HyperlinkStatusAdded)
                .ToList();
            
            var titleChanges = _document.ChangeLog.Changes
                .Where(c => c.Type == Core.Entities.ChangeType.TitleChanged || c.Type == Core.Entities.ChangeType.TitleReplaced)
                .ToList();
            
            var titleDifferences = _document.ChangeLog.Changes
                .Where(c => c.Type == Core.Entities.ChangeType.PossibleTitleChange)
                .ToList();
            
            var customHyperlinks = _document.ChangeLog.Changes
                .Where(c => c.Type == Core.Entities.ChangeType.HyperlinkUpdated && c.Description.Contains("custom", StringComparison.OrdinalIgnoreCase))
                .ToList();
            
            var customText = _document.ChangeLog.Changes
                .Where(c => c.Type == Core.Entities.ChangeType.TextReplaced)
                .ToList();
            
            var textOptimization = _document.ChangeLog.Changes
                .Where(c => c.Type == Core.Entities.ChangeType.TextOptimized)
                .ToList();

            // Create processing options
            ProcessingOptions.Add(new ProcessingOptionViewModel(
                "Updated URLs of theSource Hyperlinks",
                urlUpdates.Any(),
                urlUpdates));
            
            ProcessingOptions.Add(new ProcessingOptionViewModel(
                "Appended Content IDs to theSource Hyperlinks",
                contentIdAdditions.Any(),
                contentIdAdditions));
            
            ProcessingOptions.Add(new ProcessingOptionViewModel(
                "Checked for Expired Content",
                expiredContent.Any(),
                expiredContent));
            
            ProcessingOptions.Add(new ProcessingOptionViewModel(
                "Auto Replaced Outdated Titles",
                titleChanges.Any(),
                titleChanges));
            
            ProcessingOptions.Add(new ProcessingOptionViewModel(
                "Reported Title Differences in Changelog",
                titleDifferences.Any(),
                titleDifferences));
            
            ProcessingOptions.Add(new ProcessingOptionViewModel(
                "Replaced Custom User Defined Hyperlinks",
                customHyperlinks.Any(),
                customHyperlinks));
            
            ProcessingOptions.Add(new ProcessingOptionViewModel(
                "Replaced Custom User Defined Text",
                customText.Any(),
                customText));
            
            ProcessingOptions.Add(new ProcessingOptionViewModel(
                "Optimize Text Formatting",
                textOptimization.Any(),
                textOptimization));
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// ViewModel for individual processing options within a document
    /// </summary>
    public class ProcessingOptionViewModel : INotifyPropertyChanged
    {
        private bool _isExpanded;

        public ProcessingOptionViewModel(string optionName, bool wasProcessed, List<ChangeEntry> changes)
        {
            OptionName = optionName;
            WasProcessed = wasProcessed;
            Changes = new ObservableCollection<ProcessingChangeViewModel>(
                changes.Select(c => new ProcessingChangeViewModel(c)));
        }

        public string OptionName { get; }
        public bool WasProcessed { get; }
        public ObservableCollection<ProcessingChangeViewModel> Changes { get; }
        
        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                _isExpanded = value;
                OnPropertyChanged();
            }
        }

        public bool HasChanges => Changes.Any();
        public int ChangeCount => Changes.Count;

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// ViewModel for individual changes within a processing option
    /// </summary>
    public class ProcessingChangeViewModel
    {
        public ProcessingChangeViewModel(ChangeEntry change)
        {
            Change = change;
        }

        public ChangeEntry Change { get; }
        
        public string CurrentTitle => Change.OldValue;
        public string NuxeoTitle => ExtractNuxeoTitle();
        public string ContentId => ExtractContentId();
        public string ChangeDescription => Change.Description;
        public string ChangeType => GetChangeTypeDescription();

        private string ExtractNuxeoTitle()
        {
            // Extract Nuxeo title from change details or new value
            // This would be populated based on the actual processing logic
            return Change.NewValue ?? "N/A";
        }

        private string ExtractContentId()
        {
            // Extract content ID from element ID or details
            return Change.ElementId ?? "N/A";
        }

        private string GetChangeTypeDescription()
        {
            return Change.Type switch
            {
                Core.Entities.ChangeType.HyperlinkUpdated => "Updated URL",
                Core.Entities.ChangeType.ContentIdAdded => "Appended Content ID",
                Core.Entities.ChangeType.HyperlinkStatusAdded => "Added Status",
                Core.Entities.ChangeType.TitleChanged => "Updated Title",
                Core.Entities.ChangeType.TitleReplaced => "Replaced Title",
                Core.Entities.ChangeType.PossibleTitleChange => "Reported Title Difference",
                Core.Entities.ChangeType.TextReplaced => "Replaced Text",
                Core.Entities.ChangeType.TextOptimized => "Optimized Formatting",
                _ => Change.Type.ToString()
            };
        }
    }
}