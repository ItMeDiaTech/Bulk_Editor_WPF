using BulkEditor.Core.Interfaces;
using BulkEditor.UI.Models;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Threading;

namespace BulkEditor.UI.Services
{
    /// <summary>
    /// Implementation of notification service for UI notifications
    /// </summary>
    public class NotificationService : INotificationService
    {
        private readonly ILoggingService _logger;
        private readonly DispatcherTimer _autoHideTimer;

        public ObservableCollection<NotificationModel> Notifications { get; } = new();

        public NotificationService(ILoggingService logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));

            // Timer for auto-hiding notifications
            _autoHideTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _autoHideTimer.Tick += AutoHideTimer_Tick;
            _autoHideTimer.Start();
        }

        public void ShowInfo(string title, string message)
        {
            var notification = NotificationModel.CreateInfo(title, message);
            AddNotification(notification);
            _logger.LogInformation("Info notification: {Title} - {Message}", title, message);
        }

        public void ShowSuccess(string title, string message)
        {
            var notification = NotificationModel.CreateSuccess(title, message);
            AddNotification(notification);
            _logger.LogInformation("Success notification: {Title} - {Message}", title, message);
        }

        public void ShowWarning(string title, string message)
        {
            var notification = NotificationModel.CreateWarning(title, message);
            AddNotification(notification);
            _logger.LogWarning("Warning notification: {Title} - {Message}", title, message);
        }

        public void ShowError(string title, string message, Exception? exception = null)
        {
            var notification = NotificationModel.CreateError(title, message, exception);
            AddNotification(notification);

            if (exception != null)
            {
                _logger.LogError(exception, "Error notification: {Title} - {Message}", title, message);
            }
            else
            {
                _logger.LogError("Error notification: {Title} - {Message}", title, message);
            }
        }

        public void AddNotification(NotificationModel notification)
        {
            if (notification == null)
                throw new ArgumentNullException(nameof(notification));

            // Ensure we're on the UI thread
            if (System.Windows.Application.Current?.Dispatcher.CheckAccess() == false)
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() => AddNotification(notification));
                return;
            }

            // Limit the number of notifications (keep only the latest 10)
            while (Notifications.Count >= 10)
            {
                Notifications.RemoveAt(0);
            }

            Notifications.Add(notification);
        }

        public void RemoveNotification(NotificationModel notification)
        {
            if (notification == null)
                return;

            // Ensure we're on the UI thread
            if (System.Windows.Application.Current?.Dispatcher.CheckAccess() == false)
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() => RemoveNotification(notification));
                return;
            }

            Notifications.Remove(notification);
        }

        public void RemoveNotification(string id)
        {
            var notification = Notifications.FirstOrDefault(n => n.Id == id);
            if (notification != null)
            {
                RemoveNotification(notification);
            }
        }

        public void ClearAll()
        {
            // Ensure we're on the UI thread
            if (System.Windows.Application.Current?.Dispatcher.CheckAccess() == false)
            {
                System.Windows.Application.Current.Dispatcher.Invoke(ClearAll);
                return;
            }

            Notifications.Clear();
        }

        public void ClearBySeverity(NotificationSeverity severity)
        {
            // Ensure we're on the UI thread
            if (System.Windows.Application.Current?.Dispatcher.CheckAccess() == false)
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() => ClearBySeverity(severity));
                return;
            }

            var toRemove = Notifications.Where(n => n.Severity == severity).ToList();
            foreach (var notification in toRemove)
            {
                Notifications.Remove(notification);
            }
        }

        private void AutoHideTimer_Tick(object? sender, EventArgs e)
        {
            var now = DateTime.Now;
            var toRemove = Notifications
                .Where(n => n.AutoHide && now - n.Timestamp > n.AutoHideDelay)
                .ToList();

            foreach (var notification in toRemove)
            {
                RemoveNotification(notification);
            }
        }
    }
}