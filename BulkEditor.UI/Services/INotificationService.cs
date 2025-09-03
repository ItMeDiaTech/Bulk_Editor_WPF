using BulkEditor.UI.Models;
using System;
using System.Collections.ObjectModel;

namespace BulkEditor.UI.Services
{
    /// <summary>
    /// Service for managing application notifications
    /// </summary>
    public interface INotificationService
    {
        /// <summary>
        /// Collection of active notifications
        /// </summary>
        ObservableCollection<NotificationModel> Notifications { get; }

        /// <summary>
        /// Show an information notification
        /// </summary>
        void ShowInfo(string title, string message);

        /// <summary>
        /// Show a success notification
        /// </summary>
        void ShowSuccess(string title, string message);

        /// <summary>
        /// Show a warning notification
        /// </summary>
        void ShowWarning(string title, string message);

        /// <summary>
        /// Show an error notification
        /// </summary>
        void ShowError(string title, string message, Exception? exception = null);

        /// <summary>
        /// Add a custom notification
        /// </summary>
        void AddNotification(NotificationModel notification);

        /// <summary>
        /// Remove a specific notification
        /// </summary>
        void RemoveNotification(NotificationModel notification);

        /// <summary>
        /// Remove a notification by ID
        /// </summary>
        void RemoveNotification(string id);

        /// <summary>
        /// Clear all notifications
        /// </summary>
        void ClearAll();

        /// <summary>
        /// Clear notifications of a specific severity
        /// </summary>
        void ClearBySeverity(NotificationSeverity severity);
    }
}