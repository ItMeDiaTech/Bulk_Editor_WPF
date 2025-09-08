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

        [ObservableProperty]
        private bool _isActionable;

        [ObservableProperty]
        private string _actionText = string.Empty;

        public Action? Action { get; set; }

    // CRITICAL FIX: Cache and freeze SolidColorBrush instances to prevent TypeConverter errors
    private static readonly SolidColorBrush InfoIconBrush = CreateFrozenBrush(Color.FromRgb(0x21, 0x96, 0xF3));
    private static readonly SolidColorBrush SuccessIconBrush = CreateFrozenBrush(Color.FromRgb(0x4C, 0xAF, 0x50));
    private static readonly SolidColorBrush WarningIconBrush = CreateFrozenBrush(Color.FromRgb(0xFF, 0x98, 0x00));
    private static readonly SolidColorBrush ErrorIconBrush = CreateFrozenBrush(Color.FromRgb(0xF4, 0x43, 0x36));
    private static readonly SolidColorBrush DefaultIconBrush = CreateFrozenBrush(Color.FromRgb(0x75, 0x75, 0x75));

    private static readonly SolidColorBrush InfoBackgroundBrush = CreateFrozenBrush(Color.FromArgb(10, 0x21, 0x96, 0xF3));
    private static readonly SolidColorBrush SuccessBackgroundBrush = CreateFrozenBrush(Color.FromArgb(10, 0x4C, 0xAF, 0x50));
    private static readonly SolidColorBrush WarningBackgroundBrush = CreateFrozenBrush(Color.FromArgb(10, 0xFF, 0x98, 0x00));
    private static readonly SolidColorBrush ErrorBackgroundBrush = CreateFrozenBrush(Color.FromArgb(10, 0xF4, 0x43, 0x36));
    private static readonly SolidColorBrush DefaultBackgroundBrush = CreateFrozenBrush(Color.FromArgb(10, 0x75, 0x75, 0x75));

    private static readonly SolidColorBrush InfoBorderBrush = CreateFrozenBrush(Color.FromArgb(50, 0x21, 0x96, 0xF3));
    private static readonly SolidColorBrush SuccessBorderBrush = CreateFrozenBrush(Color.FromArgb(50, 0x4C, 0xAF, 0x50));
    private static readonly SolidColorBrush WarningBorderBrush = CreateFrozenBrush(Color.FromArgb(50, 0xFF, 0x98, 0x00));
    private static readonly SolidColorBrush ErrorBorderBrush = CreateFrozenBrush(Color.FromArgb(50, 0xF4, 0x43, 0x36));
    private static readonly SolidColorBrush DefaultBorderBrush = CreateFrozenBrush(Color.FromArgb(50, 0x75, 0x75, 0x75));

    private static readonly SolidColorBrush InfoTitleBrush = CreateFrozenBrush(Color.FromRgb(0x0D, 0x47, 0xA1));
    private static readonly SolidColorBrush SuccessTitleBrush = CreateFrozenBrush(Color.FromRgb(0x2E, 0x7D, 0x32));
    private static readonly SolidColorBrush WarningTitleBrush = CreateFrozenBrush(Color.FromRgb(0xE6, 0x5C, 0x00));
    private static readonly SolidColorBrush ErrorTitleBrush = CreateFrozenBrush(Color.FromRgb(0xC6, 0x28, 0x28));
    private static readonly SolidColorBrush DefaultTitleBrush = CreateFrozenBrush(Color.FromRgb(0x42, 0x42, 0x42));

    private static readonly SolidColorBrush MessageBrush = CreateFrozenBrush(Color.FromRgb(0x42, 0x42, 0x42));

    /// <summary>
    /// Creates a frozen SolidColorBrush for thread-safe XAML access
    /// </summary>
    private static SolidColorBrush CreateFrozenBrush(Color color)
    {
        var brush = new SolidColorBrush(color);
        brush.Freeze(); // CRITICAL: Freeze for thread-safe access
        return brush;
    }

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
            NotificationSeverity.Info => InfoIconBrush,
            NotificationSeverity.Success => SuccessIconBrush,
            NotificationSeverity.Warning => WarningIconBrush,
            NotificationSeverity.Error => ErrorIconBrush,
            _ => DefaultIconBrush
        };

        public SolidColorBrush BackgroundColor => Severity switch
        {
            NotificationSeverity.Info => InfoBackgroundBrush,
            NotificationSeverity.Success => SuccessBackgroundBrush,
            NotificationSeverity.Warning => WarningBackgroundBrush,
            NotificationSeverity.Error => ErrorBackgroundBrush,
            _ => DefaultBackgroundBrush
        };

        public SolidColorBrush BorderColor => Severity switch
        {
            NotificationSeverity.Info => InfoBorderBrush,
            NotificationSeverity.Success => SuccessBorderBrush,
            NotificationSeverity.Warning => WarningBorderBrush,
            NotificationSeverity.Error => ErrorBorderBrush,
            _ => DefaultBorderBrush
        };

        public SolidColorBrush TitleColor => Severity switch
        {
            NotificationSeverity.Info => InfoTitleBrush,
            NotificationSeverity.Success => SuccessTitleBrush,
            NotificationSeverity.Warning => WarningTitleBrush,
            NotificationSeverity.Error => ErrorTitleBrush,
            _ => DefaultTitleBrush
        };

        public SolidColorBrush MessageColor => MessageBrush;

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

        public static NotificationModel CreateActionable(string title, string message, string actionText, Action action)
        {
            return new NotificationModel
            {
                Title = title,
                Message = message,
                Severity = NotificationSeverity.Info,
                IsActionable = true,
                ActionText = actionText,
                Action = action,
                AutoHide = false
            };
        }
    }
}