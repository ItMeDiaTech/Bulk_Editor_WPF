using System;
using System.IO;
using System.Windows.Input;
using System.Windows.Media;
using BulkEditor.Core.Entities;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace BulkEditor.UI.ViewModels
{
    /// <summary>
    /// ViewModel for individual document items in the simplified document list
    /// </summary>
    public partial class DocumentListItemViewModel : ObservableObject
    {
        private readonly Document _document;
        private readonly Action<DocumentListItemViewModel> _removeDocumentAction;
        private readonly Action<DocumentListItemViewModel> _viewDetailsAction;
        private readonly Action<string> _openLocationAction;

        [ObservableProperty]
        private string _title = string.Empty;

        [ObservableProperty]
        private string _location = string.Empty;

        [ObservableProperty]
        private DocumentStatus _status;

        [ObservableProperty]
        private bool _hasErrors;

        [ObservableProperty]
        private bool _canViewDetails;

        [ObservableProperty]
        private DateTime? _processedAt;

        public DocumentListItemViewModel(
            Document document, 
            Action<DocumentListItemViewModel> removeDocumentAction,
            Action<DocumentListItemViewModel> viewDetailsAction,
            Action<string> openLocationAction)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _removeDocumentAction = removeDocumentAction ?? throw new ArgumentNullException(nameof(removeDocumentAction));
            _viewDetailsAction = viewDetailsAction ?? throw new ArgumentNullException(nameof(viewDetailsAction));
            _openLocationAction = openLocationAction ?? throw new ArgumentNullException(nameof(openLocationAction));

            UpdateFromDocument();

            // Create commands
            RemoveCommand = new RelayCommand(() => _removeDocumentAction(this));
            ViewDetailsCommand = new RelayCommand(() => _viewDetailsAction(this), () => CanViewDetails);
            OpenLocationCommand = new RelayCommand(() => _openLocationAction(Location));
        }

        public Document Document => _document;

        public ICommand RemoveCommand { get; }
        public ICommand ViewDetailsCommand { get; }
        public ICommand OpenLocationCommand { get; }

        /// <summary>
        /// Gets the status display text based on current status and error state
        /// </summary>
        public string StatusDisplayText
        {
            get
            {
                return Status switch
                {
                    DocumentStatus.Pending => "Added",
                    DocumentStatus.Processing => "Processing",
                    DocumentStatus.Completed when HasErrors => "Completed\nWith Errors",
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
                return Status switch
                {
                    DocumentStatus.Pending => new SolidColorBrush(Color.FromRgb(0x87, 0xCE, 0xFA)), // Light Blue
                    DocumentStatus.Processing => new SolidColorBrush(Color.FromRgb(0xFF, 0x98, 0x00)), // Orange
                    DocumentStatus.Completed when HasErrors => new SolidColorBrush(Color.FromRgb(0xF4, 0x43, 0x36)), // Red
                    DocumentStatus.Completed => new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50)), // Green
                    DocumentStatus.Failed => new SolidColorBrush(Color.FromRgb(0xF4, 0x43, 0x36)), // Red
                    DocumentStatus.Cancelled => new SolidColorBrush(Color.FromRgb(0x75, 0x75, 0x75)), // Gray
                    DocumentStatus.Recovered => new SolidColorBrush(Color.FromRgb(0xFF, 0x98, 0x00)), // Orange
                    _ => new SolidColorBrush(Color.FromRgb(0x75, 0x75, 0x75)) // Gray
                };
            }
        }

        /// <summary>
        /// Gets whether the status should use multiple lines (for "Completed With Errors")
        /// </summary>
        public bool IsMultiLineStatus => Status == DocumentStatus.Completed && HasErrors;

        /// <summary>
        /// Gets the top part of status text for multi-line display
        /// </summary>
        public string StatusTopText => IsMultiLineStatus ? "Completed" : StatusDisplayText;

        /// <summary>
        /// Gets the bottom part of status text for multi-line display
        /// </summary>
        public string StatusBottomText => IsMultiLineStatus ? "With Errors" : string.Empty;

        /// <summary>
        /// Updates the view model properties from the underlying document
        /// </summary>
        public void UpdateFromDocument()
        {
            Title = _document.FileName;
            Location = _document.FilePath;
            Status = _document.Status;
            HasErrors = _document.ProcessingErrors?.Count > 0;
            ProcessedAt = _document.ProcessedAt;
            
            // Can view details if document has been processed (completed, failed, or has changes)
            CanViewDetails = Status == DocumentStatus.Completed || 
                           Status == DocumentStatus.Failed || 
                           _document.ChangeLog?.Changes?.Count > 0;

            // Notify property changes for computed properties
            OnPropertyChanged(nameof(StatusDisplayText));
            OnPropertyChanged(nameof(StatusBackgroundColor));
            OnPropertyChanged(nameof(IsMultiLineStatus));
            OnPropertyChanged(nameof(StatusTopText));
            OnPropertyChanged(nameof(StatusBottomText));
            
            // Update command can execute
            ((RelayCommand)ViewDetailsCommand).NotifyCanExecuteChanged();
        }
    }
}