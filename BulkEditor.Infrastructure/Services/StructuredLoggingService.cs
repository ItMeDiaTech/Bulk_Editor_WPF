using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Implementation of structured logging service with rich HTTP and performance logging
    /// </summary>
    public class StructuredLoggingService : IStructuredLoggingService
    {
        private readonly ILoggingService _baseLogger;
        private readonly ILogger _serilogLogger;

        public StructuredLoggingService(ILoggingService baseLogger)
        {
            _baseLogger = baseLogger ?? throw new ArgumentNullException(nameof(baseLogger));
            _serilogLogger = Log.Logger;
        }

        public async Task LogHttpRequestAsync(HttpRequestLogEntry entry)
        {
            try
            {
                var properties = new Dictionary<string, object>
                {
                    ["Method"] = entry.Method,
                    ["Url"] = entry.Url,
                    ["ContentLength"] = entry.ContentLength,
                    ["CorrelationId"] = entry.CorrelationId,
                    ["OperationName"] = entry.OperationName,
                    ["Timestamp"] = entry.Timestamp,
                    ["UserAgent"] = entry.UserAgent ?? "Unknown"
                };

                // Add sanitized headers (exclude sensitive data)
                var sanitizedHeaders = SanitizeHeaders(entry.Headers);
                if (sanitizedHeaders.Any())
                {
                    properties["Headers"] = sanitizedHeaders;
                }

                // Add sanitized request body (for POST/PUT requests)
                if (!string.IsNullOrEmpty(entry.Body))
                {
                    var sanitizedBody = SanitizeRequestBody(entry.Body);
                    properties["RequestBody"] = sanitizedBody;
                }

                _serilogLogger.Information("HTTP Request: {Method} {Url} | Size: {ContentLength} bytes | Operation: {OperationName} | CorrelationId: {CorrelationId}",
                    entry.Method, entry.Url, entry.ContentLength, entry.OperationName, entry.CorrelationId);

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _baseLogger.LogError(ex, "Error logging HTTP request");
            }
        }

        public async Task LogHttpResponseAsync(HttpResponseLogEntry entry)
        {
            try
            {
                var logLevel = entry.IsSuccessStatusCode ? "Information" : "Warning";
                var statusCategory = GetStatusCodeCategory(entry.StatusCode);

                var properties = new Dictionary<string, object>
                {
                    ["StatusCode"] = entry.StatusCode,
                    ["StatusDescription"] = entry.StatusDescription,
                    ["ContentLength"] = entry.ContentLength,
                    ["Duration"] = entry.Duration.TotalMilliseconds,
                    ["CorrelationId"] = entry.CorrelationId,
                    ["OperationName"] = entry.OperationName,
                    ["IsSuccess"] = entry.IsSuccessStatusCode,
                    ["StatusCategory"] = statusCategory,
                    ["Timestamp"] = entry.Timestamp
                };

                // Add response headers (sanitized)
                var sanitizedHeaders = SanitizeHeaders(entry.Headers);
                if (sanitizedHeaders.Any())
                {
                    properties["ResponseHeaders"] = sanitizedHeaders;
                }

                // Add response body (sanitized and truncated if necessary)
                if (!string.IsNullOrEmpty(entry.Body))
                {
                    var sanitizedBody = SanitizeResponseBody(entry.Body);
                    properties["ResponseBody"] = sanitizedBody;
                }

                if (entry.IsSuccessStatusCode)
                {
                    _serilogLogger.Information("HTTP Response: {StatusCode} {StatusDescription} | Duration: {Duration}ms | Size: {ContentLength} bytes | Operation: {OperationName} | CorrelationId: {CorrelationId}",
                        entry.StatusCode, entry.StatusDescription, entry.Duration.TotalMilliseconds, entry.ContentLength, entry.OperationName, entry.CorrelationId);
                }
                else
                {
                    _serilogLogger.Warning("HTTP Response ERROR: {StatusCode} {StatusDescription} | Duration: {Duration}ms | Size: {ContentLength} bytes | Operation: {OperationName} | CorrelationId: {CorrelationId}",
                        entry.StatusCode, entry.StatusDescription, entry.Duration.TotalMilliseconds, entry.ContentLength, entry.OperationName, entry.CorrelationId);
                }

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _baseLogger.LogError(ex, "Error logging HTTP response");
            }
        }

        public void LogPerformanceMetric(PerformanceMetricEntry entry)
        {
            try
            {
                _serilogLogger.Information("Performance Metric: {OperationName} | Duration: {Duration}ms | Memory: {MemoryMB} MB | Thread: {ThreadId} | Machine: {MachineName} | CorrelationId: {CorrelationId}",
                    entry.OperationName, 
                    entry.Duration.TotalMilliseconds,
                    Math.Round(entry.MemoryUsedBytes / 1024.0 / 1024.0, 2),
                    entry.ThreadId,
                    entry.MachineName,
                    entry.CorrelationId);

                // Log additional metrics if present
                foreach (var metric in entry.Metrics)
                {
                    _serilogLogger.Debug("Metric {OperationName}.{MetricName}: {MetricValue}", 
                        entry.OperationName, metric.Key, metric.Value);
                }
            }
            catch (Exception ex)
            {
                _baseLogger.LogError(ex, "Error logging performance metric");
            }
        }

        public void LogDocumentOperation(DocumentOperationLogEntry entry)
        {
            try
            {
                var logLevel = entry.IsSuccessful ? "Information" : "Error";

                if (entry.IsSuccessful)
                {
                    _serilogLogger.Information("Document Operation: {OperationType} | File: {DocumentName} | Duration: {Duration}ms | Size: {FileSizeMB} MB | Hyperlinks: {HyperlinksProcessed} | Replacements: {TextReplacements} | CorrelationId: {CorrelationId}",
                        entry.OperationType,
                        entry.DocumentName,
                        entry.ProcessingDuration.TotalMilliseconds,
                        Math.Round(entry.FileSizeBytes / 1024.0 / 1024.0, 2),
                        entry.HyperlinksProcessed,
                        entry.TextReplacements,
                        entry.CorrelationId);
                }
                else
                {
                    _serilogLogger.Error("Document Operation FAILED: {OperationType} | File: {DocumentName} | Duration: {Duration}ms | Error: {ErrorMessage} | CorrelationId: {CorrelationId}",
                        entry.OperationType,
                        entry.DocumentName,
                        entry.ProcessingDuration.TotalMilliseconds,
                        entry.ErrorMessage ?? "Unknown error",
                        entry.CorrelationId);
                }

                // Log metadata if present
                foreach (var metadata in entry.Metadata)
                {
                    _serilogLogger.Debug("Document Metadata {DocumentName}.{MetadataKey}: {MetadataValue}",
                        entry.DocumentName, metadata.Key, metadata.Value);
                }
            }
            catch (Exception ex)
            {
                _baseLogger.LogError(ex, "Error logging document operation");
            }
        }

        public void LogBackgroundTaskEvent(BackgroundTaskLogEntry entry)
        {
            try
            {
                var logLevel = entry.EventType.ToLowerInvariant() switch
                {
                    "failed" => "Error",
                    "cancelled" => "Warning",
                    _ => "Information"
                };

                switch (logLevel)
                {
                    case "Error":
                        _serilogLogger.Error("Background Task {EventType}: {TaskName} (ID: {TaskId}) | Duration: {Duration}ms | Error: {ErrorMessage}",
                            entry.EventType, entry.TaskName, entry.TaskId, entry.Duration?.TotalMilliseconds ?? 0, entry.ErrorMessage);
                        break;
                    case "Warning":
                        _serilogLogger.Warning("Background Task {EventType}: {TaskName} (ID: {TaskId}) | Duration: {Duration}ms",
                            entry.EventType, entry.TaskName, entry.TaskId, entry.Duration?.TotalMilliseconds ?? 0);
                        break;
                    default:
                        _serilogLogger.Information("Background Task {EventType}: {TaskName} (ID: {TaskId}) | Duration: {Duration}ms",
                            entry.EventType, entry.TaskName, entry.TaskId, entry.Duration?.TotalMilliseconds ?? 0);
                        break;
                }

                // Log additional properties
                foreach (var property in entry.Properties)
                {
                    _serilogLogger.Debug("Task Property {TaskId}.{PropertyKey}: {PropertyValue}",
                        entry.TaskId, property.Key, property.Value);
                }
            }
            catch (Exception ex)
            {
                _baseLogger.LogError(ex, "Error logging background task event");
            }
        }

        public void LogRetryAttempt(RetryLogEntry entry)
        {
            try
            {
                if (entry.IsLastAttempt)
                {
                    _serilogLogger.Error("Retry EXHAUSTED: {OperationName} | Policy: {PolicyName} | Final attempt: {AttemptNumber}/{MaxAttempts} | Exception: {ExceptionType} - {ExceptionMessage} | CorrelationId: {CorrelationId}",
                        entry.OperationName, entry.PolicyName, entry.AttemptNumber, entry.MaxAttempts, entry.ExceptionType, entry.ExceptionMessage, entry.CorrelationId);
                }
                else
                {
                    _serilogLogger.Warning("Retry Attempt: {OperationName} | Policy: {PolicyName} | Attempt: {AttemptNumber}/{MaxAttempts} | Delay: {DelayMs}ms | Exception: {ExceptionType} - {ExceptionMessage} | CorrelationId: {CorrelationId}",
                        entry.OperationName, entry.PolicyName, entry.AttemptNumber, entry.MaxAttempts, entry.DelayBeforeRetry.TotalMilliseconds, entry.ExceptionType, entry.ExceptionMessage, entry.CorrelationId);
                }
            }
            catch (Exception ex)
            {
                _baseLogger.LogError(ex, "Error logging retry attempt");
            }
        }

        public IDisposable BeginScope(string operationName, Dictionary<string, object>? properties = null)
        {
            var scopeProperties = new Dictionary<string, object>
            {
                ["OperationName"] = operationName,
                ["CorrelationId"] = Guid.NewGuid().ToString(),
                ["StartTime"] = DateTime.UtcNow,
                ["ThreadId"] = Thread.CurrentThread.ManagedThreadId
            };

            if (properties != null)
            {
                foreach (var prop in properties)
                {
                    scopeProperties[prop.Key] = prop.Value;
                }
            }

            // Return a simple disposable implementation since Serilog BeginScope may not be available
            return new LogScope(operationName);
        }

        private class LogScope : IDisposable
        {
            private readonly string _operationName;
            
            public LogScope(string operationName)
            {
                _operationName = operationName;
            }

            public void Dispose()
            {
                // Cleanup logic if needed
            }
        }

        private static Dictionary<string, string> SanitizeHeaders(Dictionary<string, string> headers)
        {
            var sanitizedHeaders = new Dictionary<string, string>();
            var sensitiveHeaderNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Authorization", "Cookie", "X-API-Key", "X-Auth-Token", "Authentication"
            };

            foreach (var header in headers)
            {
                if (sensitiveHeaderNames.Contains(header.Key))
                {
                    sanitizedHeaders[header.Key] = "***REDACTED***";
                }
                else
                {
                    sanitizedHeaders[header.Key] = header.Value;
                }
            }

            return sanitizedHeaders;
        }

        private static string SanitizeRequestBody(string body)
        {
            // Truncate very large bodies
            if (body.Length > 10000)
            {
                return body.Substring(0, 10000) + "... [TRUNCATED]";
            }

            // Could add additional sanitization here for sensitive data patterns
            return body;
        }

        private static string SanitizeResponseBody(string body)
        {
            // Truncate very large bodies
            if (body.Length > 10000)
            {
                return body.Substring(0, 10000) + "... [TRUNCATED]";
            }

            return body;
        }

        private static string GetStatusCodeCategory(int statusCode)
        {
            return statusCode switch
            {
                >= 200 and < 300 => "Success",
                >= 300 and < 400 => "Redirection", 
                >= 400 and < 500 => "ClientError",
                >= 500 => "ServerError",
                _ => "Unknown"
            };
        }
    }
}