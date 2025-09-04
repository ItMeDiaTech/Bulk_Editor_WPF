using CommunityToolkit.Mvvm.ComponentModel;

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
        private bool _validateHyperlinks;

        [ObservableProperty]
        private bool _updateHyperlinks;

        [ObservableProperty]
        private bool _addContentIds;

        [ObservableProperty]
        private bool _optimizeText;

        [ObservableProperty]
        private int _timeoutPerDocumentMinutes;
    }
}