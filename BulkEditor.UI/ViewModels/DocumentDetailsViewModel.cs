using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using BulkEditor.UI.Services;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace BulkEditor.UI.ViewModels
{
    /// <summary>
    /// ViewModel for the Document Details popup window
    /// </summary>
    public partial class DocumentDetailsViewModel : ObservableObject
    {
        private readonly Document _document;
        private readonly IBackupService _backupService;
        private readonly ILoggingService _logger;
        private readonly INotificationService _notificationService;
        
        public event EventHandler? CloseRequested;

        [ObservableProperty]
        private string _documentTitle = string.Empty;

        [ObservableProperty]
        private string _documentPath = string.Empty;

        [ObservableProperty]
        private DocumentStatus _status;

        [ObservableProperty]
        private bool _hasProcessingResults;

        [ObservableProperty]
        private bool _hasBackup;

        [ObservableProperty]
        private bool _canRevert;

        [ObservableProperty]
        private DateTime? _processedAt;

        [ObservableProperty]
        private TimeSpan? _processingTime;

        public DocumentDetailsViewModel(
            Document document,
            IBackupService backupService,
            ILoggingService logger,
            INotificationService notificationService)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _backupService = backupService ?? throw new ArgumentNullException(nameof(backupService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _notificationService = notificationService ?? throw new ArgumentNullException(nameof(notificationService));

            InitializeProperties();
            InitializeProcessingOptions();

            // Initialize commands
            OpenDocumentCommand = new RelayCommand(OpenDocument, CanOpenDocument);
            OpenFileLocationCommand = new RelayCommand(OpenFileLocation);
            ViewBackupFileCommand = new RelayCommand(ViewBackupFile, () => HasBackup);
            RevertAllChangesCommand = new RelayCommand(RevertAllChanges, () => CanRevert);
            CloseCommand = new RelayCommand(Close);
        }

        public string WindowTitle => $"Document Details - {DocumentTitle}";

        public ObservableCollection<ProcessingOptionViewModel> ProcessingOptions { get; } = new();

        public ICommand OpenDocumentCommand { get; }
        public ICommand OpenFileLocationCommand { get; }
        public ICommand ViewBackupFileCommand { get; }
        public ICommand RevertAllChangesCommand { get; }
        public ICommand CloseCommand { get; }

        /// <summary>
        /// Gets the status display text based on current status and error state
        /// </summary>
        public string StatusDisplayText
        {
            get
            {
                var hasErrors = _document.ProcessingErrors?.Count > 0;
                return Status switch
                {
                    DocumentStatus.Pending => "Added",
                    DocumentStatus.Processing => "Processing",
                    DocumentStatus.Completed when hasErrors => "Completed With Errors",
                    DocumentStatus.Completed => "Completed",
                    DocumentStatus.Failed => "Failed",
                    DocumentStatus.Cancelled => "Cancelled",
                    DocumentStatus.Recovered => "Recovered",
                    _ => Status.ToString()
                };
            }
        }

        /// <summary>
        /// Gets the status background color based on current status and error state
        /// </summary>
        public SolidColorBrush StatusBackgroundColor
        {
            get
            {
                var hasErrors = _document.ProcessingErrors?.Count > 0;
                return Status switch
                {
                    DocumentStatus.Pending => new SolidColorBrush(Color.FromRgb(0x87, 0xCE, 0xFA)), // Light Blue
                    DocumentStatus.Processing => new SolidColorBrush(Color.FromRgb(0xFF, 0x98, 0x00)), // Orange
                    DocumentStatus.Completed when hasErrors => new SolidColorBrush(Color.FromRgb(0xF4, 0x43, 0x36)), // Red
                    DocumentStatus.Completed => new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50)), // Green
                    DocumentStatus.Failed => new SolidColorBrush(Color.FromRgb(0xF4, 0x43, 0x36)), // Red
                    DocumentStatus.Cancelled => new SolidColorBrush(Color.FromRgb(0x75, 0x75, 0x75)), // Gray
                    DocumentStatus.Recovered => new SolidColorBrush(Color.FromRgb(0xFF, 0x98, 0x00)), // Orange
                    _ => new SolidColorBrush(Color.FromRgb(0x75, 0x75, 0x75)) // Gray
                };
            }
        }

        public string ProcessingTimeDisplay
        {
            get
            {
                if (ProcessingTime.HasValue)
                {
                    var time = ProcessingTime.Value;
                    if (time.TotalSeconds < 1)
                        return $"{time.TotalMilliseconds:F0}ms";
                    else if (time.TotalMinutes < 1)
                        return $"{time.TotalSeconds:F1}s";
                    else
                        return $"{time.TotalMinutes:F1}m";
                }
                return "N/A";
            }
        }

        public string ProcessedAtDisplay => ProcessedAt?.ToString("yyyy-MM-dd HH:mm:ss") ?? "Not processed";

        public int TotalChangesCount => _document.ChangeLog?.Changes?.Count ?? 0;

        public int HyperlinkChangesCount => _document.ChangeLog?.Changes?
            .Count(c => c.Type == ChangeType.HyperlinkUpdated) ?? 0;

        public int TitleChangesCount => _document.ChangeLog?.Changes?
            .Count(c => c.Type == ChangeType.TitleChanged || c.Type == ChangeType.TitleReplaced) ?? 0;

        public int ErrorCount => _document.ProcessingErrors?.Count ?? 0;

        private void InitializeProperties()
        {
            DocumentTitle = _document.FileName;
            DocumentPath = _document.FilePath;
            Status = _document.Status;
            ProcessedAt = _document.ProcessedAt;
            
            // Calculate processing time if available
            if (_document.ProcessedAt.HasValue && _document.CreatedAt != default)
            {
                ProcessingTime = _document.ProcessedAt.Value - _document.CreatedAt;
            }

            HasProcessingResults = Status == DocumentStatus.Completed || 
                                 Status == DocumentStatus.Failed || 
                                 _document.ChangeLog?.Changes?.Any() == true;

            HasBackup = !string.IsNullOrEmpty(_document.BackupPath) && File.Exists(_document.BackupPath);
            CanRevert = HasBackup && (Status == DocumentStatus.Completed || Status == DocumentStatus.Failed);
        }

        private void InitializeProcessingOptions()
        {
            if (_document.ChangeLog?.Changes == null || !_document.ChangeLog.Changes.Any())
                return;

            // Group changes by processing option type (same as ProcessingResultViewModel)
            var urlUpdates = _document.ChangeLog.Changes
                .Where(c => c.Type == ChangeType.HyperlinkUpdated)
                .ToList();
            
            var contentIdAdditions = _document.ChangeLog.Changes
                .Where(c => c.Type == ChangeType.ContentIdAdded)
                .ToList();
            
            var expiredContent = _document.ChangeLog.Changes
                .Where(c => c.Type == ChangeType.HyperlinkStatusAdded)
                .ToList();
            
            var titleChanges = _document.ChangeLog.Changes
                .Where(c => c.Type == ChangeType.TitleChanged || c.Type == ChangeType.TitleReplaced)
                .ToList();
            
            var titleDifferences = _document.ChangeLog.Changes
                .Where(c => c.Type == ChangeType.PossibleTitleChange)
                .ToList();
            
            var customHyperlinks = _document.ChangeLog.Changes
                .Where(c => c.Type == ChangeType.HyperlinkUpdated && 
                           c.Description.Contains("custom", StringComparison.OrdinalIgnoreCase))
                .ToList();
            
            var customText = _document.ChangeLog.Changes
                .Where(c => c.Type == ChangeType.TextReplaced)
                .ToList();
            
            var textOptimization = _document.ChangeLog.Changes
                .Where(c => c.Type == ChangeType.TextOptimized)
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

        private bool CanOpenDocument()
        {
            return File.Exists(DocumentPath);
        }

        private void OpenDocument()
        {
            try
            {
                if (File.Exists(DocumentPath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = DocumentPath,
                        UseShellExecute = true
                    });
                    _logger.LogInformation($"Opened document: {DocumentPath}");
                }
                else
                {
                    _notificationService.ShowWarning("File Not Found", 
                        $"The document no longer exists at the specified location:\n{DocumentPath}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error opening document: {DocumentPath}", DocumentPath);
                _notificationService.ShowError("Open Document Error", 
                    "Failed to open the document.", ex);
            }
        }

        private void OpenFileLocation()
        {
            try
            {
                if (File.Exists(DocumentPath))
                {
                    // Open Windows Explorer and select the file
                    Process.Start("explorer.exe", $"/select,\"{DocumentPath}\"");
                    _logger.LogInformation($"Opened location for: {DocumentPath}");
                }
                else
                {
                    _notificationService.ShowWarning("File Not Found", 
                        $"The file no longer exists at the specified location:\n{DocumentPath}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error opening file location: {DocumentPath}", DocumentPath);
                _notificationService.ShowError("Open Location Error", 
                    "Failed to open the file location.", ex);
            }
        }

        private void ViewBackupFile()
        {
            try
            {
                if (!string.IsNullOrEmpty(_document.BackupPath) && File.Exists(_document.BackupPath))
                {
                    // Open Windows Explorer and select the backup file
                    Process.Start("explorer.exe", $"/select,\"{_document.BackupPath}\"");
                    _logger.LogInformation($"Opened backup location: {_document.BackupPath}");
                }
                else
                {
                    _notificationService.ShowWarning("Backup Not Found", 
                        "The backup file no longer exists or was not created.");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error opening backup file location: {BackupPath}", _document.BackupPath);
                _notificationService.ShowError("Open Backup Location Error", 
                    "Failed to open the backup file location.", ex);
            }
        }

        private void RevertAllChanges()
        {
            try
            {
                if (!HasBackup)
                {
                    _notificationService.ShowWarning("No Backup Available", 
                        "No backup file is available for this document.");
                    return;
                }

                // Confirm the operation
                var result = MessageBox.Show(
                    $"Are you sure you want to revert all changes to this document?\n\n" +
                    $"This will restore the backup file and overwrite the current document:\n{DocumentPath}\n\n" +
                    $"This action cannot be undone.",
                    "Confirm Revert Changes",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result != MessageBoxResult.Yes)
                    return;

                // Check if file is in use
                if (IsFileInUse(DocumentPath))
                {
                    var retryResult = MessageBox.Show(
                        $"The document appears to be open in another application.\n\n" +
                        $"Please close the document and click OK to try again, or Cancel to abort.\n\n" +
                        $"Document: {Path.GetFileName(DocumentPath)}",
                        "Document In Use",
                        MessageBoxButton.OKCancel,
                        MessageBoxImage.Warning);

                    if (retryResult == MessageBoxResult.OK)
                    {
                        // Try again
                        RevertAllChanges();
                        return;
                    }
                    else
                    {
                        return; // User cancelled
                    }
                }

                // Perform the restoration
                File.Copy(_document.BackupPath, DocumentPath, true);

                _notificationService.ShowSuccess("Revert Successful", 
                    $"Successfully reverted all changes to the document.\n\nDocument: {Path.GetFileName(DocumentPath)}");
                
                _logger.LogInformation($"Successfully reverted document: {DocumentPath} from backup: {_document.BackupPath}");
                
                MessageBox.Show(
                    "The document has been successfully reverted to its original state.",
                    "Revert Complete",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                // Close the details window
                Close();
            }
            catch (UnauthorizedAccessException ex)
            {
                _logger.LogError(ex, "Access denied when reverting document: {DocumentPath}", DocumentPath);
                _notificationService.ShowError("Access Denied", 
                    "Access denied. The document may be open in another application or you may not have sufficient permissions.", ex);
            }
            catch (IOException ex)
            {
                _logger.LogError(ex, "IO error when reverting document: {DocumentPath}", DocumentPath);
                _notificationService.ShowError("File Error", 
                    "An error occurred while accessing the file. It may be in use by another application.", ex);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error reverting document: {DocumentPath}", DocumentPath);
                _notificationService.ShowError("Revert Error", 
                    "Failed to revert the document changes.", ex);
            }
        }

        private bool IsFileInUse(string filePath)
        {
            try
            {
                using var stream = File.OpenWrite(filePath);
                return false;
            }
            catch (IOException)
            {
                return true;
            }
            catch
            {
                return false; // If we can't determine, assume it's not in use
            }
        }

        private void Close()
        {
            CloseRequested?.Invoke(this, EventArgs.Empty);
        }
    }
}