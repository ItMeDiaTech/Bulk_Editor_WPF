using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;

namespace BulkEditor.Core.Services
{
    /// <summary>
    /// Enhanced logging service with structured logging capabilities
    /// </summary>
    public interface IStructuredLoggingService
    {
        /// <summary>
        /// Logs HTTP request with detailed information
        /// </summary>
        Task LogHttpRequestAsync(HttpRequestLogEntry entry);

        /// <summary>
        /// Logs HTTP response with detailed information
        /// </summary>
        Task LogHttpResponseAsync(HttpResponseLogEntry entry);

        /// <summary>
        /// Logs performance metrics for operations
        /// </summary>
        void LogPerformanceMetric(PerformanceMetricEntry entry);

        /// <summary>
        /// Logs document processing operations with context
        /// </summary>
        void LogDocumentOperation(DocumentOperationLogEntry entry);

        /// <summary>
        /// Logs background task lifecycle events
        /// </summary>
        void LogBackgroundTaskEvent(BackgroundTaskLogEntry entry);

        /// <summary>
        /// Logs retry attempt with context
        /// </summary>
        void LogRetryAttempt(RetryLogEntry entry);

        /// <summary>
        /// Creates a scoped logging context
        /// </summary>
        IDisposable BeginScope(string operationName, Dictionary<string, object>? properties = null);
    }

    /// <summary>
    /// Structured HTTP request log entry
    /// </summary>
    public class HttpRequestLogEntry
    {
        public string Method { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
        public Dictionary<string, string> Headers { get; set; } = new();
        public string? Body { get; set; }
        public long ContentLength { get; set; }
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
        public string CorrelationId { get; set; } = string.Empty;
        public string OperationName { get; set; } = string.Empty;
        public string? UserAgent { get; set; }
    }

    /// <summary>
    /// Structured HTTP response log entry
    /// </summary>
    public class HttpResponseLogEntry
    {
        public int StatusCode { get; set; }
        public string StatusDescription { get; set; } = string.Empty;
        public Dictionary<string, string> Headers { get; set; } = new();
        public string? Body { get; set; }
        public long ContentLength { get; set; }
        public TimeSpan Duration { get; set; }
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
        public string CorrelationId { get; set; } = string.Empty;
        public string OperationName { get; set; } = string.Empty;
        public bool IsSuccessStatusCode { get; set; }
    }

    /// <summary>
    /// Performance metric log entry
    /// </summary>
    public class PerformanceMetricEntry
    {
        public string OperationName { get; set; } = string.Empty;
        public TimeSpan Duration { get; set; }
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
        public string CorrelationId { get; set; } = string.Empty;
        public Dictionary<string, object> Metrics { get; set; } = new();
        public long MemoryUsedBytes { get; set; }
        public int ThreadId { get; set; }
        public string MachineName { get; set; } = Environment.MachineName;
    }

    /// <summary>
    /// Document operation log entry
    /// </summary>
    public class DocumentOperationLogEntry
    {
        public string OperationType { get; set; } = string.Empty;
        public string DocumentPath { get; set; } = string.Empty;
        public string DocumentName { get; set; } = string.Empty;
        public long FileSizeBytes { get; set; }
        public TimeSpan ProcessingDuration { get; set; }
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
        public string CorrelationId { get; set; } = string.Empty;
        public Dictionary<string, object> Metadata { get; set; } = new();
        public int HyperlinksProcessed { get; set; }
        public int TextReplacements { get; set; }
        public bool IsSuccessful { get; set; }
        public string? ErrorMessage { get; set; }
    }

    /// <summary>
    /// Background task log entry
    /// </summary>
    public class BackgroundTaskLogEntry
    {
        public string TaskId { get; set; } = string.Empty;
        public string TaskName { get; set; } = string.Empty;
        public string EventType { get; set; } = string.Empty; // Started, Completed, Failed, Cancelled
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
        public TimeSpan? Duration { get; set; }
        public string? ErrorMessage { get; set; }
        public Dictionary<string, object> Properties { get; set; } = new();
    }

    /// <summary>
    /// Retry operation log entry
    /// </summary>
    public class RetryLogEntry
    {
        public string OperationName { get; set; } = string.Empty;
        public string PolicyName { get; set; } = string.Empty;
        public int AttemptNumber { get; set; }
        public int MaxAttempts { get; set; }
        public TimeSpan DelayBeforeRetry { get; set; }
        public string ExceptionType { get; set; } = string.Empty;
        public string ExceptionMessage { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
        public string CorrelationId { get; set; } = string.Empty;
        public bool IsLastAttempt { get; set; }
    }
}