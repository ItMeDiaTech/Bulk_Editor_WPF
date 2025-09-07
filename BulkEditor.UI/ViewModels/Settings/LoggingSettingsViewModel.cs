using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using System.Diagnostics;
using System.IO;

namespace BulkEditor.UI.ViewModels.Settings
{
    public partial class LoggingSettingsViewModel : ObservableObject
    {
        [ObservableProperty]
        private string _logLevel = string.Empty;

        [ObservableProperty]
        private string _logDirectory = string.Empty;

        [ObservableProperty]
        private bool _enableFileLogging;

        [ObservableProperty]
        private bool _enableConsoleLogging;

        [ObservableProperty]
        private int _maxLogFileSizeMB;

        [ObservableProperty]
        private int _maxLogFiles;

        [RelayCommand]
        private void BrowseLogDirectory()
        {
            var dialog = new OpenFolderDialog
            {
                Title = "Select Log Directory",
                InitialDirectory = LogDirectory
            };

            if (dialog.ShowDialog() == true)
            {
                LogDirectory = dialog.FolderName;
            }
        }

        [RelayCommand]
        private void OpenLogsFolder()
        {
            try
            {
                if (Directory.Exists(LogDirectory))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = LogDirectory,
                        UseShellExecute = true
                    });
                }
            }
            catch
            {
                // Handle error silently or show message
            }
        }
    }
}