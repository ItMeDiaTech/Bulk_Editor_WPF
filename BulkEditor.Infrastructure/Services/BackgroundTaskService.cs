using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Implementation of background task management service
    /// </summary>
    public class BackgroundTaskService : IBackgroundTaskService, IDisposable
    {
        private readonly ILoggingService _logger;
        private readonly ConcurrentDictionary<string, BackgroundTaskInfo> _tasks = new();
        private readonly ConcurrentDictionary<string, CancellationTokenSource> _cancellationTokens = new();
        private readonly object _lockObject = new object();
        private bool _disposed = false;

        public BackgroundTaskService(ILoggingService logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public event EventHandler<BackgroundTaskStatusChangedEventArgs>? TaskStatusChanged;

        public string RegisterTask(string taskName, string description = "")
        {
            if (string.IsNullOrEmpty(taskName))
                throw new ArgumentException("Task name cannot be null or empty", nameof(taskName));

            var taskId = Guid.NewGuid().ToString();
            var taskInfo = new BackgroundTaskInfo
            {
                Id = taskId,
                Name = taskName,
                Description = description,
                Status = BackgroundTaskStatus.NotStarted,
                StartTime = DateTime.Now
            };

            _tasks[taskId] = taskInfo;
            _cancellationTokens[taskId] = new CancellationTokenSource();

            _logger.LogInformation($"Registered background task: {taskName} (ID: {taskId})");
            
            OnTaskStatusChanged(taskId, BackgroundTaskStatus.NotStarted, BackgroundTaskStatus.NotStarted);
            
            return taskId;
        }

        public async Task<T> StartTaskAsync<T>(string taskId, Func<CancellationToken, Task<T>> taskFunc)
        {
            if (string.IsNullOrEmpty(taskId))
                throw new ArgumentException("Task ID cannot be null or empty", nameof(taskId));

            if (taskFunc == null)
                throw new ArgumentNullException(nameof(taskFunc));

            if (!_tasks.TryGetValue(taskId, out var taskInfo))
                throw new InvalidOperationException($"Task with ID {taskId} not found");

            if (!_cancellationTokens.TryGetValue(taskId, out var cancellationTokenSource))
                throw new InvalidOperationException($"Cancellation token for task {taskId} not found");

            var oldStatus = taskInfo.Status;
            taskInfo.Status = BackgroundTaskStatus.Running;
            taskInfo.StartTime = DateTime.Now;
            
            _logger.LogInformation($"Starting background task: {taskInfo.Name} (ID: {taskId})");
            OnTaskStatusChanged(taskId, oldStatus, BackgroundTaskStatus.Running);

            try
            {
                var result = await taskFunc(cancellationTokenSource.Token);
                
                taskInfo.Status = BackgroundTaskStatus.Completed;
                taskInfo.EndTime = DateTime.Now;
                
                _logger.LogInformation($"Background task completed successfully: {taskInfo.Name} (ID: {taskId})");
                OnTaskStatusChanged(taskId, BackgroundTaskStatus.Running, BackgroundTaskStatus.Completed);
                
                return result;
            }
            catch (OperationCanceledException ex) when (cancellationTokenSource.Token.IsCancellationRequested)
            {
                taskInfo.Status = BackgroundTaskStatus.Cancelled;
                taskInfo.EndTime = DateTime.Now;
                taskInfo.Exception = ex;
                
                _logger.LogInformation($"Background task was cancelled: {taskInfo.Name} (ID: {taskId})");
                OnTaskStatusChanged(taskId, BackgroundTaskStatus.Running, BackgroundTaskStatus.Cancelled);
                
                throw;
            }
            catch (Exception ex)
            {
                taskInfo.Status = BackgroundTaskStatus.Failed;
                taskInfo.EndTime = DateTime.Now;
                taskInfo.Exception = ex;
                
                _logger.LogError(ex, $"Background task failed: {taskInfo.Name} (ID: {taskId})");
                OnTaskStatusChanged(taskId, BackgroundTaskStatus.Running, BackgroundTaskStatus.Failed);
                
                throw;
            }
            finally
            {
                // Clean up completed/failed/cancelled tasks after a delay
                _ = Task.Delay(TimeSpan.FromMinutes(5)).ContinueWith(_ => CleanupTask(taskId));
            }
        }

        public void CancelTask(string taskId)
        {
            if (string.IsNullOrEmpty(taskId))
                return;

            if (_cancellationTokens.TryGetValue(taskId, out var cancellationTokenSource))
            {
                if (_tasks.TryGetValue(taskId, out var taskInfo))
                {
                    taskInfo.IsCancellationRequested = true;
                    _logger.LogInformation($"Cancellation requested for background task: {taskInfo.Name} (ID: {taskId})");
                }

                cancellationTokenSource.Cancel();
            }
        }

        public void CancelAllTasks()
        {
            _logger.LogInformation($"Cancelling all background tasks ({_cancellationTokens.Count} tasks)");

            var taskIds = _cancellationTokens.Keys.ToList();
            foreach (var taskId in taskIds)
            {
                CancelTask(taskId);
            }
        }

        public BackgroundTaskStatus GetTaskStatus(string taskId)
        {
            if (string.IsNullOrEmpty(taskId))
                return BackgroundTaskStatus.NotStarted;

            return _tasks.TryGetValue(taskId, out var taskInfo) 
                ? taskInfo.Status 
                : BackgroundTaskStatus.NotStarted;
        }

        public IEnumerable<BackgroundTaskInfo> GetActiveTasks()
        {
            return _tasks.Values
                .Where(task => task.Status == BackgroundTaskStatus.Running || 
                              task.Status == BackgroundTaskStatus.NotStarted)
                .ToList(); // Create a copy to avoid enumeration issues
        }

        private void OnTaskStatusChanged(string taskId, BackgroundTaskStatus oldStatus, BackgroundTaskStatus newStatus)
        {
            if (_tasks.TryGetValue(taskId, out var taskInfo))
            {
                TaskStatusChanged?.Invoke(this, new BackgroundTaskStatusChangedEventArgs
                {
                    TaskId = taskId,
                    OldStatus = oldStatus,
                    NewStatus = newStatus,
                    TaskInfo = taskInfo
                });
            }
        }

        private void CleanupTask(string taskId)
        {
            lock (_lockObject)
            {
                if (_tasks.TryGetValue(taskId, out var taskInfo) && 
                    (taskInfo.Status == BackgroundTaskStatus.Completed ||
                     taskInfo.Status == BackgroundTaskStatus.Cancelled ||
                     taskInfo.Status == BackgroundTaskStatus.Failed))
                {
                    _tasks.TryRemove(taskId, out _);
                    
                    if (_cancellationTokens.TryRemove(taskId, out var cancellationTokenSource))
                    {
                        cancellationTokenSource.Dispose();
                    }
                    
                    _logger.LogDebug($"Cleaned up background task: {taskInfo.Name} (ID: {taskId})");
                }
            }
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            _logger.LogInformation("Disposing background task service");

            // Cancel all running tasks
            CancelAllTasks();

            // Wait a short time for tasks to respond to cancellation
            Thread.Sleep(1000);

            // Dispose all cancellation token sources
            foreach (var cancellationTokenSource in _cancellationTokens.Values)
            {
                try
                {
                    cancellationTokenSource.Dispose();
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Error disposing cancellation token source: {ex.Message}");
                }
            }

            _cancellationTokens.Clear();
            _tasks.Clear();

            _disposed = true;
            GC.SuppressFinalize(this);
        }
    }
}