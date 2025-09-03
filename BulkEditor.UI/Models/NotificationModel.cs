using CommunityToolkit.Mvvm.ComponentModel;
using System;
using System.Windows.Media;

namespace BulkEditor.UI.Models
{
    /// <summary>
    /// Notification severity levels
    /// </summary>
    public enum NotificationSeverity
    {
        Info,
        Success,
        Warning,
        Error
    }

    /// <summary>
    /// Notification model for UI display
    /// </summary>
    public partial class NotificationModel : ObservableObject
    {
        [ObservableProperty]
        private string _id = Guid.NewGuid().ToString();

        [ObservableProperty]
        private string _title = string.Empty;

        [ObservableProperty]
        private string _message = string.Empty;

        [ObservableProperty]
        private NotificationSeverity _severity = NotificationSeverity.Info;

        [ObservableProperty]
        private DateTime _timestamp = DateTime.Now;

        [ObservableProperty]
        private bool _autoHide = true;

        [ObservableProperty]
        private TimeSpan _autoHideDelay = TimeSpan.FromSeconds(5);

        [ObservableProperty]
        private Exception? _exception;

        public string Icon => Severity switch
        {
            NotificationSeverity.Info => "ℹ",
            NotificationSeverity.Success => "✓",
            NotificationSeverity.Warning => "⚠",
            NotificationSeverity.Error => "✕",
            _ => "ℹ"
        };

        public SolidColorBrush IconColor => Severity switch
        {
            NotificationSeverity.Info => new SolidColorBrush(Color.FromRgb(0x21, 0x96, 0xF3)),
            NotificationSeverity.Success => new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50)),
            NotificationSeverity.Warning => new SolidColorBrush(Color.FromRgb(0xFF, 0x98, 0x00)),
            NotificationSeverity.Error => new SolidColorBrush(Color.FromRgb(0xF4, 0x43, 0x36)),
            _ => new SolidColorBrush(Color.FromRgb(0x75, 0x75, 0x75))
        };

        public SolidColorBrush BackgroundColor => Severity switch
        {
            NotificationSeverity.Info => new SolidColorBrush(Color.FromArgb(10, 0x21, 0x96, 0xF3)),
            NotificationSeverity.Success => new SolidColorBrush(Color.FromArgb(10, 0x4C, 0xAF, 0x50)),
            NotificationSeverity.Warning => new SolidColorBrush(Color.FromArgb(10, 0xFF, 0x98, 0x00)),
            NotificationSeverity.Error => new SolidColorBrush(Color.FromArgb(10, 0xF4, 0x43, 0x36)),
            _ => new SolidColorBrush(Color.FromArgb(10, 0x75, 0x75, 0x75))
        };

        public SolidColorBrush BorderColor => Severity switch
        {
            NotificationSeverity.Info => new SolidColorBrush(Color.FromArgb(50, 0x21, 0x96, 0xF3)),
            NotificationSeverity.Success => new SolidColorBrush(Color.FromArgb(50, 0x4C, 0xAF, 0x50)),
            NotificationSeverity.Warning => new SolidColorBrush(Color.FromArgb(50, 0xFF, 0x98, 0x00)),
            NotificationSeverity.Error => new SolidColorBrush(Color.FromArgb(50, 0xF4, 0x43, 0x36)),
            _ => new SolidColorBrush(Color.FromArgb(50, 0x75, 0x75, 0x75))
        };

        public SolidColorBrush TitleColor => Severity switch
        {
            NotificationSeverity.Info => new SolidColorBrush(Color.FromRgb(0x0D, 0x47, 0xA1)),
            NotificationSeverity.Success => new SolidColorBrush(Color.FromRgb(0x2E, 0x7D, 0x32)),
            NotificationSeverity.Warning => new SolidColorBrush(Color.FromRgb(0xE6, 0x5C, 0x00)),
            NotificationSeverity.Error => new SolidColorBrush(Color.FromRgb(0xC6, 0x28, 0x28)),
            _ => new SolidColorBrush(Color.FromRgb(0x42, 0x42, 0x42))
        };

        public SolidColorBrush MessageColor => new SolidColorBrush(Color.FromRgb(0x42, 0x42, 0x42));

        public static NotificationModel CreateInfo(string title, string message)
        {
            return new NotificationModel
            {
                Title = title,
                Message = message,
                Severity = NotificationSeverity.Info
            };
        }

        public static NotificationModel CreateSuccess(string title, string message)
        {
            return new NotificationModel
            {
                Title = title,
                Message = message,
                Severity = NotificationSeverity.Success
            };
        }

        public static NotificationModel CreateWarning(string title, string message)
        {
            return new NotificationModel
            {
                Title = title,
                Message = message,
                Severity = NotificationSeverity.Warning,
                AutoHideDelay = TimeSpan.FromSeconds(8)
            };
        }

        public static NotificationModel CreateError(string title, string message, Exception? exception = null)
        {
            return new NotificationModel
            {
                Title = title,
                Message = message,
                Severity = NotificationSeverity.Error,
                Exception = exception,
                AutoHide = false // Error notifications should be manually dismissed
            };
        }
    }
}