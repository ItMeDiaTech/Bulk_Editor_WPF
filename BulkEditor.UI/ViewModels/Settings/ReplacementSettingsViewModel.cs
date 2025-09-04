using System.Collections.ObjectModel;
using BulkEditor.Core.Configuration;
using CommunityToolkit.Mvvm.ComponentModel;

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
    }
}