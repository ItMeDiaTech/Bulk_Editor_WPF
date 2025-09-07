using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Core.Services
{
    /// <summary>
    /// Service for managing background tasks with cancellation support
    /// </summary>
    public interface IBackgroundTaskService
    {
        /// <summary>
        /// Registers a new background task
        /// </summary>
        string RegisterTask(string taskName, string description = "");

        /// <summary>
        /// Starts a background task with cancellation support
        /// </summary>
        Task<T> StartTaskAsync<T>(string taskId, Func<CancellationToken, Task<T>> taskFunc);

        /// <summary>
        /// Cancels a specific task
        /// </summary>
        void CancelTask(string taskId);

        /// <summary>
        /// Cancels all running tasks
        /// </summary>
        void CancelAllTasks();

        /// <summary>
        /// Gets the status of a specific task
        /// </summary>
        BackgroundTaskStatus GetTaskStatus(string taskId);

        /// <summary>
        /// Gets all active tasks
        /// </summary>
        IEnumerable<BackgroundTaskInfo> GetActiveTasks();

        /// <summary>
        /// Event raised when a task status changes
        /// </summary>
        event EventHandler<BackgroundTaskStatusChangedEventArgs> TaskStatusChanged;

        /// <summary>
        /// Disposes all resources and cancels all tasks
        /// </summary>
        void Dispose();
    }

    /// <summary>
    /// Represents the status of a background task
    /// </summary>
    public enum BackgroundTaskStatus
    {
        NotStarted,
        Running,
        Completed,
        Cancelled,
        Failed
    }

    /// <summary>
    /// Information about a background task
    /// </summary>
    public class BackgroundTaskInfo
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public BackgroundTaskStatus Status { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime? EndTime { get; set; }
        public TimeSpan ElapsedTime => EndTime?.Subtract(StartTime) ?? DateTime.Now.Subtract(StartTime);
        public Exception? Exception { get; set; }
        public bool IsCancellationRequested { get; set; }
    }

    /// <summary>
    /// Event arguments for task status changes
    /// </summary>
    public class BackgroundTaskStatusChangedEventArgs : EventArgs
    {
        public string TaskId { get; set; } = string.Empty;
        public BackgroundTaskStatus OldStatus { get; set; }
        public BackgroundTaskStatus NewStatus { get; set; }
        public BackgroundTaskInfo TaskInfo { get; set; } = null!;
    }
}