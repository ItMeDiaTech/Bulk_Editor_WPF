using System;
using System.Collections.Generic;

namespace BulkEditor.Core.Models
{
    /// <summary>
    /// Represents performance metrics for an operation
    /// </summary>
    public class PerformanceMetric
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public string OperationName { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
        public TimeSpan Duration { get; set; }
        public long MemoryUsedBytes { get; set; }
        public int ThreadId { get; set; }
        public string MachineName { get; set; } = string.Empty;
        public Dictionary<string, object> CustomMetrics { get; set; } = new();
    }
}