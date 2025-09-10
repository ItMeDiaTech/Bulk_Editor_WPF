using BulkEditor.Application.Services;
using BulkEditor.Core.Configuration;
using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using BulkEditor.UI.Services;
using BulkEditor.UI.Views;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Timers;

namespace BulkEditor.UI.ViewModels
{
    /// <summary>
    /// Main window ViewModel for the Bulk Editor application
    /// </summary>
    public partial class MainWindowViewModel : ViewModelBase
    {
        private readonly IApplicationService _applicationService;
        private readonly ILoggingService _logger;
        private readonly INotificationService _notificationService;
        private readonly IServiceProvider _serviceProvider;
        private readonly AppSettings _appSettings;
        private readonly IUndoService _undoService;
        private readonly ISessionManager _sessionManager;
        private readonly IBackupService _backupService;
        private readonly IBackgroundTaskService _backgroundTaskService;
        private readonly BulkEditor.Application.Services.UpdateManager _updateManager;
        private string? _currentTaskId;

        // CRITICAL FIX: Add freeze detection and prevention
        private readonly System.Timers.Timer _freezeDetectionTimer;
        private DateTime _lastUIUpdateTime = DateTime.UtcNow;
        private bool _freezeDetectionEnabled = true;
        
        // CRITICAL FIX: Add real-time timer updates for processing
        private readonly System.Timers.Timer _processingTimer;
        private BatchProcessingProgress? _currentBatchProgress;

        [ObservableProperty]
        private ObservableCollection<DocumentViewModel> _documents = new();

        [ObservableProperty]
        private ObservableCollection<ProcessingResultViewModel> _processingResults = new();

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasDocuments))]
        private ObservableCollection<DocumentListItemViewModel> _documentItems = new();

        private ObservableCollection<DocumentListItemViewModel> _allDocumentItems = new();

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasSearchText))]
        private string _searchText = string.Empty;

        public bool HasSearchText => !string.IsNullOrEmpty(SearchText);

        partial void OnSearchTextChanged(string value)
        {
            FilterDocuments();
        }

        private void FilterDocuments()
        {
            try
            {
                // CRITICAL FIX: Ensure we're on the UI thread when modifying ObservableCollection
                if (!System.Windows.Application.Current?.Dispatcher?.CheckAccess() == true)
                {
                    System.Windows.Application.Current?.Dispatcher?.Invoke(() => FilterDocuments());
                    return;
                }

                DocumentItems.Clear();

                if (string.IsNullOrEmpty(SearchText))
                {
                    // Show all documents when search is empty
                    foreach (var item in _allDocumentItems)
                    {
                        DocumentItems.Add(item);
                    }
                }
                else
                {
                    // Filter documents by title or content ID
                    var searchTerm = SearchText.ToLowerInvariant();
                    var filteredItems = _allDocumentItems.Where(item =>
                        item.Title.ToLowerInvariant().Contains(searchTerm) ||
                        item.Location.ToLowerInvariant().Contains(searchTerm)
                    );

                    foreach (var item in filteredItems)
                    {
                        DocumentItems.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error filtering documents with search term: {SearchTerm}", SearchText);
            }
        }

        [ObservableProperty]
        private DocumentViewModel? _selectedDocument;

        [ObservableProperty]
        private string _statusMessage = "Ready";

        [ObservableProperty]
        private double _progressValue;

        [ObservableProperty]
        private string _progressMessage = string.Empty;

        [ObservableProperty]
        private bool _isProcessing = false;

        [ObservableProperty]
        private string _statusIcon = "ℹ";

        [ObservableProperty]
        private bool _showStatusIcon = true;

        [ObservableProperty]
        private SolidColorBrush _statusBackgroundColor = new(Color.FromRgb(0x21, 0x96, 0xF3));

        [ObservableProperty]
        private SolidColorBrush _statusIconColor = new(Color.FromRgb(0x21, 0x96, 0xF3));

        [ObservableProperty]
        private SolidColorBrush _statusTextColor = new(Colors.White);

        [ObservableProperty]
        private ProcessingStatistics? _processingStatistics;

        [ObservableProperty]
        private int _totalDocuments;

        [ObservableProperty]
        private int _processedDocuments;

        [ObservableProperty]
        private int _failedDocuments;

        [ObservableProperty]
        private bool _isRevertEnabled;

        [ObservableProperty]
        private bool _isBusy = false;

        [ObservableProperty]
        private string _busyMessage = "Loading...";

        [ObservableProperty]
        private string _currentOperation = string.Empty;

        [ObservableProperty]
        private string _elapsedTime = "00:00:00";

        [ObservableProperty]
        private string _estimatedTimeRemaining = "Calculating...";

        [ObservableProperty]
        private double _currentDocumentProgress = 0;

        [ObservableProperty]
        private int _totalHyperlinksFound = 0;

        [ObservableProperty]
        private int _totalHyperlinksProcessed = 0;

        [ObservableProperty]
        private int _totalHyperlinksUpdated = 0;

        [ObservableProperty]
        private int _totalTextReplacements = 0;

        [ObservableProperty]
        private string _averageProcessingTime = "0.00s";

        [ObservableProperty]
        private ObservableCollection<string> _recentErrors = new();

        [ObservableProperty]
        private ObservableCollection<BackgroundTaskInfo> _activeTasks = new();

        [ObservableProperty]
        private double _overallProgress = 0;

        [RelayCommand]
        private void ClearSearch()
        {
            SearchText = string.Empty;
        }

        public bool HasProcessingResults => ProcessingResults.Any();

        public bool HasDocuments => DocumentItems.Any();

        public INotificationService NotificationService => _notificationService;

        public MainWindowViewModel(IApplicationService applicationService, ILoggingService logger, INotificationService notificationService, IServiceProvider serviceProvider, AppSettings appSettings, IUndoService undoService, ISessionManager sessionManager, IBackupService backupService, IBackgroundTaskService backgroundTaskService, BulkEditor.Application.Services.UpdateManager updateManager)
        {
            _applicationService = applicationService ?? throw new ArgumentNullException(nameof(applicationService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _notificationService = notificationService ?? throw new ArgumentNullException(nameof(notificationService));
            _serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));
            _appSettings = appSettings ?? throw new ArgumentNullException(nameof(appSettings));
            _undoService = undoService;
            _sessionManager = sessionManager;
            _backupService = backupService;
            _backgroundTaskService = backgroundTaskService ?? throw new ArgumentNullException(nameof(backgroundTaskService));
            _updateManager = updateManager ?? throw new ArgumentNullException(nameof(updateManager));

            Title = "Bulk Document Editor";
            SetStatusReady();

            // CRITICAL FIX: Initialize freeze detection timer with reasonable intervals
            _freezeDetectionTimer = new System.Timers.Timer(10000); // Check every 10 seconds (less aggressive)
            _freezeDetectionTimer.Elapsed += OnFreezeDetectionCheck;
            _freezeDetectionTimer.AutoReset = true;
            _freezeDetectionTimer.Start();

            _logger.LogInformation("Freeze detection timer started - monitoring UI responsiveness");
            
            // CRITICAL FIX: Initialize real-time processing timer for elapsed time updates
            _processingTimer = new System.Timers.Timer(1000); // Update every 1 second
            _processingTimer.Elapsed += OnProcessingTimerElapsed;
            _processingTimer.AutoReset = true;
            // Timer will be started when processing begins

            // Subscribe to background task status changes
            _backgroundTaskService.TaskStatusChanged += OnBackgroundTaskStatusChanged;

            // CRITICAL FIX: Subscribe to update manager events for automatic notifications
            _updateManager.UpdateAvailable += OnUpdateAvailable;

            // Subscribe to collection changes to update command states
            Documents.CollectionChanged += (s, e) =>
            {
                UpdateUIHeartbeat(); // Update freeze detection
                ProcessDocumentsCommand.NotifyCanExecuteChanged();
                OnPropertyChanged(nameof(Documents));
            };

            ProcessingResults.CollectionChanged += (s, e) =>
            {
                UpdateUIHeartbeat(); // Update freeze detection
                OnPropertyChanged(nameof(HasProcessingResults));
            };

            DocumentItems.CollectionChanged += (s, e) =>
            {
                UpdateUIHeartbeat(); // Update freeze detection
                OnPropertyChanged(nameof(HasDocuments));
                ClearDocumentsCommand.NotifyCanExecuteChanged();
            };
        }

        private void SetStatusReady()
        {
            StatusMessage = "Ready - Select Word documents to process";
            StatusIcon = "ℹ";
            StatusBackgroundColor = new SolidColorBrush(Color.FromRgb(0x21, 0x96, 0xF3));
            StatusIconColor = new SolidColorBrush(Color.FromRgb(0x21, 0x96, 0xF3));
            StatusTextColor = new SolidColorBrush(Colors.White);
            ShowStatusIcon = true;
        }

        private void SetStatusSuccess(string message)
        {
            StatusMessage = message;
            StatusIcon = "✓";
            StatusBackgroundColor = new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50));
            StatusIconColor = new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50));
            StatusTextColor = new SolidColorBrush(Colors.White);
            ShowStatusIcon = true;
        }

        private void SetStatusWarning(string message)
        {
            StatusMessage = message;
            StatusIcon = "⚠";
            StatusBackgroundColor = new SolidColorBrush(Color.FromRgb(0xFF, 0x98, 0x00));
            StatusIconColor = new SolidColorBrush(Color.FromRgb(0xFF, 0x98, 0x00));
            StatusTextColor = new SolidColorBrush(Colors.White);
            ShowStatusIcon = true;
        }

        private void SetStatusError(string message)
        {
            StatusMessage = message;
            StatusIcon = "✕";
            StatusBackgroundColor = new SolidColorBrush(Color.FromRgb(0xF4, 0x43, 0x36));
            StatusIconColor = new SolidColorBrush(Color.FromRgb(0xF4, 0x43, 0x36));
            StatusTextColor = new SolidColorBrush(Colors.White);
            ShowStatusIcon = true;
        }

        private void SetStatusProcessing(string message)
        {
            StatusMessage = message;
            StatusIcon = "⚙";
            StatusBackgroundColor = new SolidColorBrush(Color.FromRgb(0x21, 0x96, 0xF3));
            StatusIconColor = new SolidColorBrush(Color.FromRgb(0x21, 0x96, 0xF3));
            StatusTextColor = new SolidColorBrush(Colors.White);
            ShowStatusIcon = true;
        }

        [RelayCommand]
        private async Task SelectFilesAsync()
        {
            try
            {
                _logger.LogInformation("Starting file selection dialog");

                var openFileDialog = new OpenFileDialog
                {
                    Title = "Select Word Documents",
                    Filter = "Word Documents (*.docx;*.docm)|*.docx;*.docm|All Files (*.*)|*.*",
                    Multiselect = true,
                    CheckFileExists = true
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    _logger.LogInformation("Files selected: {FileCount}", openFileDialog.FileNames.Length);
                    await AddFilesAsync(openFileDialog.FileNames);
                }
                else
                {
                    _logger.LogInformation("File selection was cancelled by user");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "CRITICAL ERROR in SelectFilesAsync: {ErrorMessage}", ex.Message);
                _logger.LogError("Stack trace: {StackTrace}", ex.StackTrace ?? "No stack trace available");
                _logger.LogError("Thread ID: {ThreadId}", System.Threading.Thread.CurrentThread.ManagedThreadId);

                try
                {
                    _notificationService.ShowError("File Selection Error", "Failed to select files for processing. Please try again or restart the application.", ex);
                    SetStatusError($"Error selecting files: {ex.Message}");
                }
                catch (Exception innerEx)
                {
                    _logger.LogError(innerEx, "CRITICAL: Error in error handling for SelectFilesAsync");
                }
            }
        }

        [RelayCommand]
        private async Task AddFilesAsync(string[] filePaths)
        {
            try
            {
                _logger.LogInformation("Starting file addition process for {Count} files", filePaths?.Length ?? 0);

                if (filePaths == null || filePaths.Length == 0)
                {
                    _logger.LogWarning("No files provided for addition");
                    SetStatusError("No files were provided");
                    return;
                }

                SetStatusProcessing("Validating selected files...");

                // Add timeout protection for file validation to prevent hanging on network drives
                using var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(30));

                var validation = await _applicationService.ValidateFilesAsync(filePaths, timeoutCts.Token);
                _logger.LogInformation("Validation completed: {ValidCount} valid, {InvalidCount} invalid files",
                    validation.ValidFiles.Count, validation.InvalidFiles.Count);

                SetStatusProcessing($"Adding {validation.ValidFiles.Count} valid documents...");

                var processedCount = 0;
                foreach (var filePath in validation.ValidFiles)
                {
                    if (!Documents.Any(d => d.FilePath == filePath))
                    {
                        // Create Document entity with proper initialization
                        var document = new Document
                        {
                            FilePath = filePath,
                            FileName = Path.GetFileName(filePath),
                            Status = DocumentStatus.Pending,
                            ChangeLog = new ChangeLog(),
                            ProcessingErrors = new List<ProcessingError>(),
                            CreatedAt = DateTime.Now
                        };

                        // Add to original Documents collection
                        Documents.Add(new DocumentViewModel
                        {
                            FilePath = filePath,
                            FileName = Path.GetFileName(filePath),
                            Status = DocumentStatus.Pending,
                            Document = document
                        });

                        // Update progress
                        processedCount++;
                        SetStatusProcessing($"Adding documents... ({processedCount}/{validation.ValidFiles.Count})");

                        // Add to new DocumentItems collection for display with comprehensive error handling
                        try
                        {
                            _logger.LogInformation("Creating DocumentListItemViewModel for: {FilePath}", filePath);

                            // CRITICAL FIX: Wrap DocumentListItemViewModel creation in additional safety
                            DocumentListItemViewModel? documentItem = null;
                            try
                            {
                                documentItem = new DocumentListItemViewModel(
                                    document,
                                    RemoveDocumentFromList,
                                    ViewDocumentDetails,
                                    OpenDocumentLocation
                                );
                                _logger.LogInformation("Successfully created DocumentListItemViewModel for: {FilePath}", filePath);
                            }
                            catch (Exception createEx)
                            {
                                _logger.LogError(createEx, "Failed to create DocumentListItemViewModel for: {FilePath}", filePath);
                                throw new InvalidOperationException($"Failed to create document list item for {filePath}", createEx);
                            }

                            // CRITICAL FIX: Ensure thread-safe collection access with additional validation
                            if (documentItem != null)
                            {
                                try
                                {
                                    if (System.Windows.Application.Current?.Dispatcher != null)
                                    {
                                        await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                                        {
                                            try
                                            {
                                                _allDocumentItems.Add(documentItem);
                                                _logger.LogInformation("Added DocumentListItemViewModel to collection: {FilePath}", filePath);
                                            }
                                            catch (Exception addEx)
                                            {
                                                _logger.LogError(addEx, "Failed to add document item to collection: {FilePath}", filePath);
                                                throw;
                                            }
                                        });
                                    }
                                    else
                                    {
                                        _allDocumentItems.Add(documentItem);
                                        _logger.LogInformation("Added DocumentListItemViewModel to collection (no dispatcher): {FilePath}", filePath);
                                    }
                                }
                                catch (Exception collectionEx)
                                {
                                    _logger.LogError(collectionEx, "Failed to add DocumentListItemViewModel to collection: {FilePath}", filePath);
                                    throw new InvalidOperationException($"Failed to add document to collection: {filePath}", collectionEx);
                                }
                            }
                            else
                            {
                                _logger.LogError("DocumentListItemViewModel is null for: {FilePath}", filePath);
                                throw new InvalidOperationException($"Document item creation returned null for {filePath}");
                            }
                        }
                        catch (Exception docItemEx)
                        {
                            _logger.LogError(docItemEx, "CRITICAL: Complete failure creating/adding DocumentListItemViewModel for file: {FilePath}", filePath);
                            // Don't rethrow here - continue with other files but log the failure
                            _notificationService.ShowWarning("Document Display Error",
                                $"Failed to add {Path.GetFileName(filePath)} to the document list. It has been added to processing but may not display correctly.");
                        }
                    }
                }

                // CRITICAL FIX: Apply filter to refresh DocumentItems display after adding files
                // Ensure this runs on the UI thread to prevent crashes
                try
                {
                    _logger.LogInformation("Applying FilterDocuments to refresh UI display");
                    if (System.Windows.Application.Current?.Dispatcher != null)
                    {
                        await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                        {
                            try
                            {
                                FilterDocuments();
                                _logger.LogInformation("Successfully applied FilterDocuments via Dispatcher");
                            }
                            catch (Exception filterEx)
                            {
                                _logger.LogError(filterEx, "Error in FilterDocuments via Dispatcher");
                                throw;
                            }
                        });
                    }
                    else
                    {
                        FilterDocuments();
                        _logger.LogInformation("Successfully applied FilterDocuments without Dispatcher");
                    }
                }
                catch (Exception filterDocEx)
                {
                    _logger.LogError(filterDocEx, "CRITICAL: Failed to apply FilterDocuments - documents may not display");
                    _notificationService.ShowWarning("Display Update Error",
                        "Documents were added successfully but may not appear in the list immediately. Try refreshing or restarting the application.");
                }

                if (validation.InvalidFiles.Any())
                {
                    var invalidFiles = string.Join(", ", validation.InvalidFiles.Take(3));
                    if (validation.InvalidFiles.Count > 3)
                        invalidFiles += $" and {validation.InvalidFiles.Count - 3} more";

                    _notificationService.ShowWarning("Invalid Files Detected",
                        $"Some files were skipped: {invalidFiles}");

                    SetStatusWarning($"Some files skipped - {Documents.Count} documents ready for processing");
                }
                else
                {
                    var newFilesCount = Documents.Count;
                    TotalDocuments = newFilesCount;

                    if (newFilesCount > 0)
                    {
                        _notificationService.ShowSuccess("Files Added",
                            $"Successfully added {newFilesCount} document(s) for processing.");
                        SetStatusSuccess($"Ready - {newFilesCount} documents selected for processing");
                    }
                    else
                    {
                        SetStatusReady();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding files: {ErrorMessage}", ex.Message);
                _logger.LogError(ex, "Stack trace: {StackTrace}", ex.StackTrace ?? "No stack trace available");

                // Additional logging for debugging thread safety issues
                _logger.LogError("Application dispatcher availability: {DispatcherAvailable}",
                    System.Windows.Application.Current?.Dispatcher != null);
                _logger.LogError("Current thread ID: {ThreadId}", System.Threading.Thread.CurrentThread.ManagedThreadId);

                _notificationService.ShowError("File Validation Error", "Failed to validate selected files. Check logs for details.", ex);
                SetStatusError($"Error adding files: {ex.Message}");
            }
        }

        [RelayCommand(CanExecute = nameof(CanProcess))]
        private async Task ProcessDocumentsAsync()
        {
            if (IsProcessing) return;

            // Register background task
            _currentTaskId = _backgroundTaskService.RegisterTask(
                "Document Processing",
                $"Processing {Documents.Count} documents");

            try
            {
                IsProcessing = true;
                ProcessedDocuments = 0;
                FailedDocuments = 0;
                ProgressValue = 0;

                // Start background task
                var results = await _backgroundTaskService.StartTaskAsync(_currentTaskId, async (cancellationToken) =>
                {
                    // Start a new session for this processing job
                    var session = _sessionManager.StartSession();
                    _logger.LogInformation("Processing session {SessionId} started.", session.SessionId);
                    IsRevertEnabled = false; // Disable revert until processing is complete

                    SetStatusProcessing("Validating files before processing...");
                    var filePaths = Documents.Select(d => d.FilePath).ToList();
                    
                    // Validate that all files exist before processing
                    var validFilePaths = new List<string>();
                    var missingFiles = new List<string>();
                    
                    foreach (var filePath in filePaths)
                    {
                        if (File.Exists(filePath))
                        {
                            validFilePaths.Add(filePath);
                        }
                        else
                        {
                            missingFiles.Add(filePath);
                            _logger.LogWarning("File not found during validation: {FilePath}", filePath);
                        }
                    }
                    
                    // Remove missing files from the UI collection
                    if (missingFiles.Count > 0)
                    {
                        await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                        {
                            var documentsToRemove = Documents.Where(d => missingFiles.Contains(d.FilePath)).ToList();
                            foreach (var doc in documentsToRemove)
                            {
                                Documents.Remove(doc);
                                _logger.LogInformation("Removed missing file from UI: {FilePath}", doc.FilePath);
                            }
                        });
                        
                        _logger.LogWarning("Removed {Count} missing files from processing queue", missingFiles.Count);
                    }
                    
                    // If no valid files remain, abort processing
                    if (validFilePaths.Count == 0)
                    {
                        throw new InvalidOperationException("No valid files found for processing. All files appear to be missing.");
                    }
                    
                    SetStatusProcessing("Backing up files before processing...");

                    // Create backups before processing (with individual file error handling)
                    var validFilesWithBackups = new List<string>();
                    foreach (var filePath in validFilePaths)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        
                        try
                        {
                            // Double-check file exists immediately before backup
                            if (!File.Exists(filePath))
                            {
                                _logger.LogWarning("File disappeared before backup: {FilePath}", filePath);
                                continue;
                            }
                            
                            var backupPath = await _backupService.CreateBackupAsync(filePath, session);
                            _sessionManager.AddFileToSession(filePath, backupPath);
                            validFilesWithBackups.Add(filePath);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "Failed to create backup for {FilePath}, skipping file", filePath);
                            // Continue with other files instead of failing entire batch
                        }
                    }
                    
                    // If no files have backups, abort processing
                    if (validFilesWithBackups.Count == 0)
                    {
                        throw new InvalidOperationException("No files could be backed up. Processing aborted.");
                    }
                    
                    // Update valid file paths to only include files that were successfully backed up
                    validFilePaths = validFilesWithBackups;

                    SetStatusProcessing("Starting document processing...");

                    var progress = new Progress<BatchProcessingProgress>(OnBatchProgressChanged);

                    return await _applicationService.ProcessDocumentsBatchAsync(
                        validFilePaths,
                        progress,
                        cancellationToken);
                });

                // Update document ViewModels with results
                foreach (var result in results)
                {
                    var documentVM = Documents.FirstOrDefault(d => d.FilePath == result.FilePath);
                    if (documentVM != null)
                    {
                        UpdateDocumentViewModel(documentVM, result);
                    }
                }

                // Update ProcessingResults for the modern tree view
                ProcessingResults.Clear();
                foreach (var result in results)
                {
                    ProcessingResults.Add(new ProcessingResultViewModel(result));
                }

                ProcessingStatistics = _applicationService.GetProcessingStatistics(results);
                
                // CRITICAL FIX: Stop real-time timer when processing completes
                IsProcessing = false;
                _processingTimer.Stop();
                _currentBatchProgress = null;

                ProgressValue = 100;

                // Show completion notification and status
                if (ProcessingStatistics.FailedDocuments == 0)
                {
                    _notificationService.ShowSuccess("Processing Complete",
                        $"Successfully processed all {ProcessingStatistics.SuccessfulDocuments} documents.");
                    SetStatusSuccess($"Processing completed - All {ProcessingStatistics.SuccessfulDocuments} documents processed successfully");
                }
                else
                {
                    _notificationService.ShowWarning("Processing Complete with Errors",
                        $"Processed {ProcessingStatistics.SuccessfulDocuments} documents successfully, {ProcessingStatistics.FailedDocuments} failed.");
                    SetStatusWarning($"Processing completed - {ProcessingStatistics.SuccessfulDocuments} successful, {ProcessingStatistics.FailedDocuments} failed");
                }

                _logger.LogInformation("Batch processing completed successfully");
                IsRevertEnabled = _undoService.CanUndo();
            }
            catch (OperationCanceledException)
            {
                SetStatusWarning("Processing cancelled by user");
                _logger.LogInformation("Batch processing cancelled by user");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during batch processing");
                _notificationService.ShowError("Processing Failed", "An error occurred during document processing.", ex);
                SetStatusError($"Processing failed: {ex.Message}");
            }
            finally
            {
                IsProcessing = false;
                _currentTaskId = null;
                IsRevertEnabled = _undoService.CanUndo();
            }
        }

        private bool CanProcess() => Documents.Any() && !IsProcessing;

        [RelayCommand(CanExecute = nameof(CanCancel))]
        private void CancelProcessing()
        {
            if (!string.IsNullOrEmpty(_currentTaskId))
            {
                _backgroundTaskService.CancelTask(_currentTaskId);
                SetStatusWarning("Cancelling processing...");
                _logger.LogInformation("User requested cancellation of background task: {TaskId}", _currentTaskId);
            }
        }

        private bool CanCancel() => IsProcessing && !string.IsNullOrEmpty(_currentTaskId);

        [RelayCommand]
        private void ClearDocuments()
        {
            if (IsProcessing)
            {
                MessageBox.Show("Cannot clear documents while processing is in progress.", "Operation Not Allowed",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Documents.Clear();
            DocumentItems.Clear();
            _allDocumentItems.Clear(); // CRITICAL FIX: Also clear the master collection to prevent removed files from reappearing
            ProcessingResults.Clear();
            SelectedDocument = null;
            TotalDocuments = 0;
            ProcessedDocuments = 0;
            FailedDocuments = 0;
            ProgressValue = 0;
            ProcessingStatistics = null;
            
            // CRITICAL FIX: Stop real-time timer when resetting
            IsProcessing = false;
            _processingTimer.Stop();
            _currentBatchProgress = null;
            
            SetStatusReady();
            IsRevertEnabled = _undoService.CanUndo();
        }

        [RelayCommand(CanExecute = nameof(CanRevert))]
        private async Task RevertLastSessionAsync()
        {
            if (MessageBox.Show("Are you sure you want to revert the last processing session? This will restore all processed files to their original state and cannot be undone.", "Confirm Revert", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }

            SetStatusProcessing("Reverting last session...");

            var success = await _undoService.UndoLastSessionAsync();

            if (success)
            {
                _notificationService.ShowSuccess("Revert Complete", "Successfully reverted the last processing session.");
                SetStatusSuccess("Last processing session reverted successfully.");
                // Reset UI to a clean state
                ClearDocuments();
            }
            else
            {
                _notificationService.ShowError("Revert Failed", "Failed to revert one or more files. Please check the logs for more details.");
                SetStatusError("Revert failed. Check logs for details.");
            }

            IsRevertEnabled = _undoService.CanUndo();
        }

        private bool CanRevert() => IsRevertEnabled && !IsProcessing;

        [RelayCommand]
        private async Task ExportResultsAsync()
        {
            try
            {
                if (ProcessingStatistics == null || !Documents.Any())
                {
                    MessageBox.Show("No processing results to export.", "Export Results",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var saveFileDialog = new SaveFileDialog
                {
                    Title = "Export Processing Results",
                    Filter = "JSON Files (*.json)|*.json|CSV Files (*.csv)|*.csv",
                    DefaultExt = "json",
                    FileName = $"BulkEditor_Results_{DateTime.Now:yyyyMMdd_HHmmss}"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    var format = Path.GetExtension(saveFileDialog.FileName).ToLower() switch
                    {
                        ".csv" => ExportFormat.Csv,
                        ".json" => ExportFormat.Json,
                        _ => ExportFormat.Json
                    };

                    var documents = Documents.Where(d => d.Document != null).Select(d => d.Document!).ToList();

                    var success = await _applicationService.ExportResultsAsync(
                        documents,
                        saveFileDialog.FileName,
                        format);

                    if (success)
                    {
                        _notificationService.ShowSuccess("Export Complete",
                            $"Results successfully exported to {Path.GetFileName(saveFileDialog.FileName)}");
                        SetStatusSuccess($"Results exported to {Path.GetFileName(saveFileDialog.FileName)}");
                    }
                    else
                    {
                        _notificationService.ShowError("Export Failed",
                            "Failed to export results. Check the application logs for details.");
                        SetStatusError("Export failed - Check logs for details");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting results");
                _notificationService.ShowError("Export Error", "Failed to export processing results.", ex);
                SetStatusError($"Export error: {ex.Message}");
            }
        }

        private void OnBatchProgressChanged(BatchProcessingProgress progress)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                // CRITICAL FIX: Store progress for real-time timer updates
                _currentBatchProgress = progress;
                
                // Start timer on first progress update
                if (!IsProcessing)
                {
                    IsProcessing = true;
                    _processingTimer.Start();
                }

                // Update basic progress
                TotalDocuments = progress.TotalDocuments;
                ProcessedDocuments = progress.ProcessedDocuments;
                FailedDocuments = progress.FailedDocuments;
                ProgressValue = progress.PercentageComplete;
                OverallProgress = progress.OverallProgress;

                // Update detailed progress information
                CurrentOperation = progress.CurrentOperation;
                CurrentDocumentProgress = progress.CurrentDocumentProgress;
                ElapsedTime = progress.FormattedElapsedTime;
                EstimatedTimeRemaining = progress.FormattedEstimatedTimeRemaining;

                // Update processing statistics
                TotalHyperlinksFound = progress.TotalHyperlinksFound;
                TotalHyperlinksProcessed = progress.TotalHyperlinksProcessed;
                TotalHyperlinksUpdated = progress.TotalHyperlinksUpdated;
                TotalTextReplacements = progress.TotalTextReplacements;
                AverageProcessingTime = $"{progress.AverageProcessingTimePerDocument:F2}s";

                // Update recent errors (keep only last 5)
                RecentErrors.Clear();
                foreach (var error in progress.RecentErrors.TakeLast(5))
                {
                    RecentErrors.Add(error);
                }

                // Update progress message with more detail
                var currentDoc = !string.IsNullOrEmpty(progress.CurrentDocument)
                    ? Path.GetFileName(progress.CurrentDocument)
                    : "Unknown";

                ProgressMessage = string.IsNullOrEmpty(progress.CurrentOperation)
                    ? $"Processing: {currentDoc}"
                    : $"{progress.CurrentOperation}: {currentDoc}";

                SetStatusProcessing($"Processing {progress.ProcessedDocuments}/{progress.TotalDocuments} documents - {progress.FormattedElapsedTime} elapsed");
            });
        }

        private void UpdateDocumentViewModel(DocumentViewModel documentVM, Document result)
        {
            documentVM.Document = result;
            documentVM.Status = result.Status;
            documentVM.ProcessedAt = result.ProcessedAt;
            documentVM.HyperlinkCount = result.Hyperlinks.Count;
            documentVM.UpdatedHyperlinks = result.Hyperlinks.Count(h => h.ActionTaken == HyperlinkAction.Updated);
            documentVM.ErrorCount = result.ProcessingErrors.Count;
            documentVM.HasErrors = result.ProcessingErrors.Any();

            // Also update the corresponding DocumentListItemViewModel
            var documentListItem = DocumentItems.FirstOrDefault(item => item.Document.Id == result.Id);
            if (documentListItem != null)
            {
                // Update the document reference and refresh the view model
                documentListItem.Document.Status = result.Status;
                documentListItem.Document.ProcessedAt = result.ProcessedAt;
                documentListItem.Document.ProcessingErrors = result.ProcessingErrors;
                documentListItem.Document.ChangeLog = result.ChangeLog;
                documentListItem.Document.BackupPath = result.BackupPath;

                // Trigger property updates on the list item
                documentListItem.UpdateFromDocument();
            }
        }

        partial void OnSelectedDocumentChanged(DocumentViewModel? value)
        {
            // Handle selection changes if needed
        }

        public override async Task InitializeAsync()
        {
            await base.InitializeAsync();
            _logger.LogInformation("MainWindowViewModel initialized");
            _notificationService.ShowInfo("Application Ready", "Bulk Document Editor is ready for use.");
        }

        private void OnBackgroundTaskStatusChanged(object? sender, BackgroundTaskStatusChangedEventArgs e)
        {
            try
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    // Update active tasks collection
                    var activeTasks = _backgroundTaskService.GetActiveTasks().ToList();
                    ActiveTasks.Clear();
                    foreach (var task in activeTasks)
                    {
                        ActiveTasks.Add(task);
                    }

                    // Update command states
                    CancelProcessingCommand.NotifyCanExecuteChanged();
                    ProcessDocumentsCommand.NotifyCanExecuteChanged();

                    _logger.LogDebug("Background task {TaskId} status changed: {OldStatus} -> {NewStatus}",
                        e.TaskId, e.OldStatus, e.NewStatus);
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error handling background task status change");
            }
        }

        public override void Cleanup()
        {
            // CRITICAL FIX: Cleanup freeze detection timer
            _freezeDetectionTimer?.Stop();
            _freezeDetectionTimer?.Dispose();

            // Unsubscribe from events
            _backgroundTaskService.TaskStatusChanged -= OnBackgroundTaskStatusChanged;
            _updateManager.UpdateAvailable -= OnUpdateAvailable;

            // Cancel any running background tasks
            if (!string.IsNullOrEmpty(_currentTaskId))
            {
                _backgroundTaskService.CancelTask(_currentTaskId);
            }
            _backgroundTaskService.CancelAllTasks();

            base.Cleanup();
        }

        [RelayCommand]
        private void CloseNotification(object parameter)
        {
            if (parameter is BulkEditor.UI.Models.NotificationModel notification)
            {
                _notificationService.RemoveNotification(notification);
            }
        }

        [RelayCommand]
        private void OpenSettings()
        {
            try
            {
                var settingsViewModel = _serviceProvider.GetRequiredService<SettingsViewModel>();
                var settingsWindow = new SettingsWindow(settingsViewModel);

                settingsWindow.Owner = System.Windows.Application.Current.MainWindow;

                var result = settingsWindow.ShowDialog();

                if (result == true)
                {
                    _notificationService.ShowSuccess("Settings Saved", "Application settings have been updated successfully.");
                    SetStatusSuccess("Settings updated successfully");
                    _logger.LogInformation("Settings updated by user");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error opening settings window");
                _notificationService.ShowError("Settings Error", "Failed to open settings window.", ex);
                SetStatusError("Failed to open settings");
            }
        }

        #region Document List Actions

        /// <summary>
        /// Removes a document from the processing list
        /// </summary>
        private void RemoveDocumentFromList(DocumentListItemViewModel documentItem)
        {
            try
            {
                // Remove from both collections
                var documentViewModel = Documents.FirstOrDefault(d => d.Document?.Id == documentItem.Document.Id);
                if (documentViewModel != null)
                {
                    Documents.Remove(documentViewModel);
                }

                DocumentItems.Remove(documentItem);
                _allDocumentItems.Remove(documentItem); // CRITICAL FIX: Also remove from master collection to prevent it from reappearing

                // Update counts
                TotalDocuments = Documents.Count;

                // Update status
                if (Documents.Count == 0)
                {
                    SetStatusReady();
                }
                else
                {
                    SetStatusSuccess($"Ready - {Documents.Count} documents selected for processing");
                }

                _logger.LogInformation($"Document removed from list: {documentItem.Document.FileName}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error removing document from list");
                _notificationService.ShowError("Remove Error", "Failed to remove document from list.", ex);
            }
        }

        /// <summary>
        /// Opens the document details popup window
        /// </summary>
        private void ViewDocumentDetails(DocumentListItemViewModel documentItem)
        {
            try
            {
                var detailsViewModel = new DocumentDetailsViewModel(
                    documentItem.Document,
                    _backupService,
                    _logger,
                    _notificationService);

                var detailsWindow = new Views.DocumentDetailsWindow(detailsViewModel)
                {
                    Owner = System.Windows.Application.Current.MainWindow
                };

                detailsWindow.ShowDialog();

                _logger.LogInformation($"Viewed details for document: {documentItem.Document.FileName}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error viewing document details");
                _notificationService.ShowError("View Details Error", "Failed to open document details.", ex);
            }
        }

        /// <summary>
        /// Opens the dedicated processing options window
        /// </summary>
        /// <summary>
        /// CRITICAL FIX: Async method to open processing settings without blocking UI thread
        /// </summary>
        [RelayCommand]
        private async Task OpenProcessingSettingsAsync()
        {
            try
            {
                _logger.LogDebug("Opening processing options window asynchronously");

                // CRITICAL FIX: Create ViewModel without synchronous loading
                var processingOptionsViewModel = new SimpleProcessingOptionsViewModel(
                    _logger,
                    _notificationService,
                    _serviceProvider.GetRequiredService<BulkEditor.Core.Services.IConfigurationService>());

                var processingOptionsWindow = new Views.ProcessingOptionsWindow(processingOptionsViewModel)
                {
                    Owner = System.Windows.Application.Current.MainWindow
                };

                // Show dialog asynchronously to prevent blocking
                var result = processingOptionsWindow.ShowDialog();

                if (result == true)
                {
                    _notificationService.ShowSuccess("Processing Options Saved", "Processing options have been updated successfully.");
                    SetStatusSuccess("Processing options updated successfully");
                    _logger.LogInformation("Processing options updated by user");
                }

                _logger.LogInformation("Processing options window opened successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error opening processing options window asynchronously");
                _notificationService.ShowError("Processing Options Error", "Failed to open processing options window.", ex);
                SetStatusError("Failed to open processing options");
            }
        }

        /// <summary>
        /// Opens the file location in Windows Explorer
        /// </summary>
        private void OpenDocumentLocation(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    // Open Windows Explorer and select the file
                    System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{filePath}\"");
                    _logger.LogInformation($"Opened location for: {filePath}");
                }
                else
                {
                    _notificationService.ShowWarning("File Not Found",
                        $"The file no longer exists at the specified location:\n{filePath}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error opening document location");
                _notificationService.ShowError("Open Location Error", "Failed to open document location.", ex);
            }
        }

        /// <summary>
        /// Handle update availability notifications and show them to the user
        /// </summary>
        private void OnUpdateAvailable(object? sender, UpdateAvailableEventArgs e)
        {
            try
            {
                _logger.LogInformation("Update available: Version {Version}", e.UpdateInfo.Version);

                // Show notification to user about available update
                var currentVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "Unknown";
                var message = $"A new version ({e.UpdateInfo.Version}) of BulkEditor is available!\n\n" +
                              $"Current version: {currentVersion}\n" +
                              $"Release notes: {e.UpdateInfo.ReleaseNotes ?? "See GitHub for details"}\n\n" +
                              "Would you like to download and install the update now?";

                var result = System.Windows.MessageBox.Show(
                    message,
                    "Update Available",
                    System.Windows.MessageBoxButton.YesNo,
                    System.Windows.MessageBoxImage.Information);

                if (result == System.Windows.MessageBoxResult.Yes)
                {
                    // User wants to install the update
                    _logger.LogInformation("User chose to install update {Version}", e.UpdateInfo.Version);
                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            var updateService = _serviceProvider.GetRequiredService<IUpdateService>();
                            await updateService.DownloadAndInstallUpdateAsync(e.UpdateInfo);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "Failed to download and install update");
                            await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                            {
                                _notificationService.ShowError("Update Error", "Failed to download and install update. Please try again later.", ex);
                            });
                        }
                    });
                }
                else
                {
                    _logger.LogInformation("User declined to install update {Version}", e.UpdateInfo.Version);
                    _notificationService.ShowSuccess("Update Available", "A new version is available. You can check for updates later in Settings.");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error handling update available notification");
                _notificationService.ShowError("Update Notification Error", "Failed to process update notification.", ex);
            }
        }

        #endregion

        #region Freeze Detection and Prevention

        /// <summary>
        /// CRITICAL FIX: Updates the UI heartbeat to indicate the UI is responsive
        /// </summary>
        private void UpdateUIHeartbeat()
        {
            _lastUIUpdateTime = DateTime.UtcNow;
        }

        /// <summary>
        /// CRITICAL FIX: Monitors for UI freeze conditions
        /// </summary>
        private void OnFreezeDetectionCheck(object? sender, System.Timers.ElapsedEventArgs e)
        {
            if (!_freezeDetectionEnabled) return;

            try
            {
                var timeSinceLastUpdate = DateTime.UtcNow - _lastUIUpdateTime;

                // CRITICAL FIX: More reasonable freeze detection timeout (90 seconds instead of 30)
                // This allows for normal HTTP timeouts (60s) plus buffer for processing
                if (timeSinceLastUpdate.TotalSeconds > 90)
                {
                    _logger.LogWarning("FREEZE DETECTION: UI appears unresponsive for {Seconds} seconds",
                        timeSinceLastUpdate.TotalSeconds);

                    // Only intervene if freeze persists for more than 2 minutes
                    if (timeSinceLastUpdate.TotalSeconds > 120 && IsProcessing)
                    {
                        _logger.LogError("SEVERE FREEZE DETECTED after 2+ minutes - attempting recovery");

                        // Try to cancel current operation as last resort
                        try
                        {
                            if (!string.IsNullOrEmpty(_currentTaskId))
                            {
                                _backgroundTaskService.CancelTask(_currentTaskId);
                                _logger.LogInformation("Cancelled background task {TaskId} due to severe freeze detection", _currentTaskId);

                                // Show user notification about the recovery attempt
                                System.Windows.Application.Current?.Dispatcher?.BeginInvoke(() =>
                                {
                                    _notificationService.ShowWarning("Processing Recovery",
                                        "The application detected an unresponsive operation and attempted automatic recovery. " +
                                        "If this continues, please restart the application.");
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "Failed to cancel background task during freeze recovery");
                        }
                    }

                    // Reset heartbeat to prevent continuous warnings
                    UpdateUIHeartbeat();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in freeze detection check");
            }
        }

        /// <summary>
        /// CRITICAL FIX: Timer handler for real-time elapsed time updates during processing
        /// </summary>
        private void OnProcessingTimerElapsed(object? sender, System.Timers.ElapsedEventArgs e)
        {
            if (!IsProcessing || _currentBatchProgress == null)
                return;

            try
            {
                // Update the elapsed time display in real-time
                System.Windows.Application.Current?.Dispatcher?.BeginInvoke(() =>
                {
                    if (_currentBatchProgress != null)
                    {
                        ElapsedTime = _currentBatchProgress.FormattedElapsedTime;
                        EstimatedTimeRemaining = _currentBatchProgress.FormattedEstimatedTimeRemaining;
                        
                        // Update status with current elapsed time
                        if (_currentBatchProgress.ProcessedDocuments > 0)
                        {
                            SetStatusProcessing($"Processing {_currentBatchProgress.ProcessedDocuments}/{_currentBatchProgress.TotalDocuments} documents - {_currentBatchProgress.FormattedElapsedTime} elapsed");
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating processing timer display");
            }
        }

        /// <summary>
        /// CRITICAL FIX: Wraps async operations with timeout and freeze prevention
        /// </summary>
        private async Task<T> ExecuteWithTimeoutAsync<T>(Func<CancellationToken, Task<T>> operation,
            TimeSpan timeout, string operationName)
        {
            UpdateUIHeartbeat();

            using var cts = new CancellationTokenSource(timeout);

            try
            {
                _logger.LogDebug("Starting operation {OperationName} with timeout {Timeout}",
                    operationName, timeout);

                var result = await operation(cts.Token).ConfigureAwait(false);

                _logger.LogDebug("Completed operation {OperationName} successfully", operationName);
                UpdateUIHeartbeat();

                return result;
            }
            catch (OperationCanceledException) when (cts.Token.IsCancellationRequested)
            {
                _logger.LogError("Operation {OperationName} timed out after {Timeout}", operationName, timeout);
                throw new TimeoutException($"Operation '{operationName}' timed out after {timeout}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Operation {OperationName} failed with error", operationName);
                throw;
            }
        }

        #endregion
    }

    /// <summary>
    /// ViewModel for individual documents in the list
    /// </summary>
    public partial class DocumentViewModel : ObservableObject
    {
        [ObservableProperty]
        private string _filePath = string.Empty;

        [ObservableProperty]
        private string _fileName = string.Empty;

        [ObservableProperty]
        private DocumentStatus _status = DocumentStatus.Pending;

        [ObservableProperty]
        private DateTime? _processedAt;

        [ObservableProperty]
        private int _hyperlinkCount;

        [ObservableProperty]
        private int _updatedHyperlinks;

        [ObservableProperty]
        private int _errorCount;

        [ObservableProperty]
        private bool _hasErrors;

        public Document? Document { get; set; }

        public string StatusDisplay => Status.ToString();

        public string ProcessedAtDisplay => ProcessedAt?.ToString("yyyy-MM-dd HH:mm:ss") ?? "Not processed";

        public string SummaryDisplay => Document?.ChangeLog.Summary ?? "No changes";
    }
}