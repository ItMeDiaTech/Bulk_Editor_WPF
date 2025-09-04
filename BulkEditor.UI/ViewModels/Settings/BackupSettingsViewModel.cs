using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;

namespace BulkEditor.UI.ViewModels.Settings
{
    public partial class BackupSettingsViewModel : ObservableObject
    {
        [ObservableProperty]
        private string _backupDirectory = string.Empty;

        [ObservableProperty]
        private bool _createTimestampedBackups;

        [ObservableProperty]
        private bool _compressBackups;

        [ObservableProperty]
        private bool _autoCleanupOldBackups;

        [ObservableProperty]
        private int _maxBackupAge;

        [RelayCommand]
        private void BrowseBackupDirectory()
        {
            var dialog = new OpenFolderDialog
            {
                Title = "Select Backup Directory",
                InitialDirectory = BackupDirectory
            };

            if (dialog.ShowDialog() == true)
            {
                BackupDirectory = dialog.FolderName;
            }
        }
    }
}