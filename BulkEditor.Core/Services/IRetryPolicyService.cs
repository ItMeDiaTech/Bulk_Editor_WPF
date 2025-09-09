using System;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Core.Services
{
    /// <summary>
    /// Service for intelligent retry logic with various policies
    /// </summary>
    public interface IRetryPolicyService
    {
        /// <summary>
        /// Executes an operation with exponential backoff retry policy
        /// </summary>
        Task<T> ExecuteWithRetryAsync<T>(
            Func<Task<T>> operation,
            RetryPolicy policy,
            CancellationToken cancellationToken = default);

        /// <summary>
        /// Executes an operation with exponential backoff retry policy (void return)
        /// </summary>
        Task ExecuteWithRetryAsync(
            Func<Task> operation,
            RetryPolicy policy,
            CancellationToken cancellationToken = default);

        /// <summary>
        /// Creates a retry policy for HTTP operations
        /// </summary>
        RetryPolicy CreateHttpRetryPolicy();

        /// <summary>
        /// Creates a retry policy for file operations
        /// </summary>
        RetryPolicy CreateFileRetryPolicy();

        /// <summary>
        /// Creates a retry policy for OpenXML operations
        /// </summary>
        RetryPolicy CreateOpenXmlRetryPolicy();

        /// <summary>
        /// Creates a retry policy for database operations
        /// </summary>
        RetryPolicy CreateDatabaseRetryPolicy();

        /// <summary>
        /// Creates a custom retry policy
        /// </summary>
        RetryPolicy CreateCustomRetryPolicy(
            int maxRetries,
            TimeSpan baseDelay,
            RetryBackoffType backoffType = RetryBackoffType.Exponential,
            Func<Exception, bool>? shouldRetry = null);
    }

    /// <summary>
    /// Retry policy configuration
    /// </summary>
    public class RetryPolicy
    {
        public int MaxRetries { get; set; } = 3;
        public TimeSpan BaseDelay { get; set; } = TimeSpan.FromMilliseconds(100);
        public TimeSpan MaxDelay { get; set; } = TimeSpan.FromSeconds(10);
        public RetryBackoffType BackoffType { get; set; } = RetryBackoffType.Exponential;
        public double BackoffMultiplier { get; set; } = 2.0;
        public double JitterMaxPercent { get; set; } = 0.1; // 10% jitter
        public Func<Exception, bool>? ShouldRetry { get; set; }
        public string PolicyName { get; set; } = "Default";
    }

    /// <summary>
    /// Types of backoff strategies for retry logic
    /// </summary>
    public enum RetryBackoffType
    {
        /// <summary>
        /// Fixed delay between retries
        /// </summary>
        Fixed,

        /// <summary>
        /// Linear increase in delay (baseDelay * attempt)
        /// </summary>
        Linear,

        /// <summary>
        /// Exponential increase in delay (baseDelay * 2^attempt)
        /// </summary>
        Exponential,

        /// <summary>
        /// Exponential with jitter to avoid thundering herd
        /// </summary>
        ExponentialWithJitter
    }

    /// <summary>
    /// Context information about retry attempts
    /// </summary>
    public class RetryContext
    {
        public int AttemptNumber { get; set; }
        public int MaxRetries { get; set; }
        public TimeSpan ElapsedTime { get; set; }
        public Exception? LastException { get; set; }
        public string PolicyName { get; set; } = string.Empty;
        public bool IsLastAttempt => AttemptNumber >= MaxRetries;
    }

    /// <summary>
    /// Event arguments for retry events
    /// </summary>
    public class RetryEventArgs : EventArgs
    {
        public RetryContext Context { get; set; } = null!;
        public TimeSpan DelayBeforeNextAttempt { get; set; }
    }
}