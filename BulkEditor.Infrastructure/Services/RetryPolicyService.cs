using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Implementation of intelligent retry policies for various operation types
    /// </summary>
    public class RetryPolicyService : IRetryPolicyService
    {
        private readonly ILoggingService _logger;
        private readonly Random _random = new();

        public RetryPolicyService(ILoggingService logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public async Task<T> ExecuteWithRetryAsync<T>(
            Func<Task<T>> operation,
            RetryPolicy policy,
            CancellationToken cancellationToken = default)
        {
            if (operation == null) throw new ArgumentNullException(nameof(operation));
            if (policy == null) throw new ArgumentNullException(nameof(policy));

            var startTime = DateTime.UtcNow;
            Exception? lastException = null;

            for (int attempt = 1; attempt <= policy.MaxRetries + 1; attempt++)
            {
                var context = new RetryContext
                {
                    AttemptNumber = attempt,
                    MaxRetries = policy.MaxRetries + 1,
                    ElapsedTime = DateTime.UtcNow - startTime,
                    LastException = lastException,
                    PolicyName = policy.PolicyName
                };

                try
                {
                    if (attempt > 1)
                    {
                        _logger.LogInformation("Retry attempt {Attempt}/{Max} for policy '{Policy}'", 
                            attempt, context.MaxRetries, policy.PolicyName);
                    }

                    var result = await operation().ConfigureAwait(false);
                    
                    if (attempt > 1)
                    {
                        _logger.LogInformation("Operation succeeded after {Attempts} attempts using policy '{Policy}'", 
                            attempt, policy.PolicyName);
                    }

                    return result;
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    context.LastException = ex;

                    // Check if we should retry this exception
                    if (!ShouldRetryException(ex, policy))
                    {
                        _logger.LogError(ex, "Operation failed with non-retryable exception using policy '{Policy}'", policy.PolicyName);
                        throw;
                    }

                    // If this is the last attempt, throw
                    if (context.IsLastAttempt)
                    {
                        _logger.LogError(ex, "Operation failed after {Attempts} attempts using policy '{Policy}'", 
                            context.MaxRetries, policy.PolicyName);
                        throw new RetryExhaustedException(
                            $"Operation failed after {context.MaxRetries} attempts using policy '{policy.PolicyName}'", ex);
                    }

                    // Calculate delay before next attempt
                    var delay = CalculateDelay(attempt - 1, policy);
                    
                    _logger.LogWarning("Operation failed on attempt {Attempt}/{Max}, retrying in {Delay}ms. Error: {Error}", 
                        attempt, context.MaxRetries, delay.TotalMilliseconds, ex.Message);

                    // Wait before next attempt
                    await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
                }
            }

            // This should never be reached, but just in case
            throw new InvalidOperationException("Retry logic error - this should not be reached");
        }

        public async Task ExecuteWithRetryAsync(
            Func<Task> operation,
            RetryPolicy policy,
            CancellationToken cancellationToken = default)
        {
            await ExecuteWithRetryAsync(async () =>
            {
                await operation().ConfigureAwait(false);
                return true; // Dummy return value
            }, policy, cancellationToken).ConfigureAwait(false);
        }

        public RetryPolicy CreateHttpRetryPolicy()
        {
            return new RetryPolicy
            {
                MaxRetries = 3,
                BaseDelay = TimeSpan.FromMilliseconds(500),
                MaxDelay = TimeSpan.FromSeconds(30),
                BackoffType = RetryBackoffType.ExponentialWithJitter,
                BackoffMultiplier = 2.0,
                JitterMaxPercent = 0.2,
                PolicyName = "HTTP",
                ShouldRetry = IsRetryableHttpException
            };
        }

        public RetryPolicy CreateFileRetryPolicy()
        {
            return new RetryPolicy
            {
                MaxRetries = 5,
                BaseDelay = TimeSpan.FromMilliseconds(100),
                MaxDelay = TimeSpan.FromSeconds(5),
                BackoffType = RetryBackoffType.Exponential,
                BackoffMultiplier = 1.5,
                JitterMaxPercent = 0.1,
                PolicyName = "File",
                ShouldRetry = IsRetryableFileException
            };
        }

        public RetryPolicy CreateOpenXmlRetryPolicy()
        {
            return new RetryPolicy
            {
                MaxRetries = 3,
                BaseDelay = TimeSpan.FromMilliseconds(200),
                MaxDelay = TimeSpan.FromSeconds(2),
                BackoffType = RetryBackoffType.Linear,
                BackoffMultiplier = 1.0,
                JitterMaxPercent = 0.05,
                PolicyName = "OpenXML",
                ShouldRetry = IsRetryableOpenXmlException
            };
        }

        public RetryPolicy CreateDatabaseRetryPolicy()
        {
            return new RetryPolicy
            {
                MaxRetries = 3,
                BaseDelay = TimeSpan.FromMilliseconds(250),
                MaxDelay = TimeSpan.FromSeconds(10),
                BackoffType = RetryBackoffType.ExponentialWithJitter,
                BackoffMultiplier = 2.5,
                JitterMaxPercent = 0.3,
                PolicyName = "Database",
                ShouldRetry = IsRetryableDatabaseException
            };
        }

        public RetryPolicy CreateCustomRetryPolicy(
            int maxRetries,
            TimeSpan baseDelay,
            RetryBackoffType backoffType = RetryBackoffType.Exponential,
            Func<Exception, bool>? shouldRetry = null)
        {
            return new RetryPolicy
            {
                MaxRetries = maxRetries,
                BaseDelay = baseDelay,
                BackoffType = backoffType,
                PolicyName = "Custom",
                ShouldRetry = shouldRetry ?? (_ => true)
            };
        }

        private bool ShouldRetryException(Exception ex, RetryPolicy policy)
        {
            if (policy.ShouldRetry != null)
                return policy.ShouldRetry(ex);

            // Default retry logic for common transient exceptions
            return IsRetryableException(ex);
        }

        private static bool IsRetryableException(Exception ex)
        {
            return ex switch
            {
                TimeoutException => true,
                HttpRequestException => true,
                IOException when ex.Message.Contains("being used by another process") => true,
                IOException when ex.Message.Contains("sharing violation") => true,
                UnauthorizedAccessException => true,
                _ => false
            };
        }

        private static bool IsRetryableHttpException(Exception ex)
        {
            return ex switch
            {
                HttpRequestException httpEx => true,
                TaskCanceledException when ex.Message.Contains("timeout") => true,
                TimeoutException => true,
                WebException webEx => webEx.Status switch
                {
                    WebExceptionStatus.Timeout => true,
                    WebExceptionStatus.ConnectionClosed => true,
                    WebExceptionStatus.ConnectFailure => true,
                    WebExceptionStatus.NameResolutionFailure => false, // DNS issues shouldn't be retried immediately
                    WebExceptionStatus.ProxyNameResolutionFailure => false,
                    _ => true
                },
                _ => false
            };
        }

        private static bool IsRetryableFileException(Exception ex)
        {
            return ex switch
            {
                IOException ioEx when ioEx.Message.Contains("being used by another process") => true,
                IOException ioEx when ioEx.Message.Contains("sharing violation") => true,
                UnauthorizedAccessException => true,
                DirectoryNotFoundException => false, // Path issues shouldn't be retried
                FileNotFoundException => false,
                _ => false
            };
        }

        private static bool IsRetryableOpenXmlException(Exception ex)
        {
            return ex switch
            {
                IOException ioEx when ioEx.Message.Contains("being used by another process") => true,
                IOException ioEx when ioEx.Message.Contains("sharing violation") => true,
                UnauthorizedAccessException => true,
                _ => false
            };
        }

        private static bool IsRetryableDatabaseException(Exception ex)
        {
            // This is a placeholder for database-specific exception handling
            // Would be expanded based on actual database provider used
            return ex switch
            {
                TimeoutException => true,
                InvalidOperationException when ex.Message.Contains("timeout") => true,
                _ => false
            };
        }

        private TimeSpan CalculateDelay(int attemptNumber, RetryPolicy policy)
        {
            TimeSpan delay = policy.BackoffType switch
            {
                RetryBackoffType.Fixed => policy.BaseDelay,
                RetryBackoffType.Linear => TimeSpan.FromMilliseconds(policy.BaseDelay.TotalMilliseconds * (attemptNumber + 1)),
                RetryBackoffType.Exponential => TimeSpan.FromMilliseconds(policy.BaseDelay.TotalMilliseconds * Math.Pow(policy.BackoffMultiplier, attemptNumber)),
                RetryBackoffType.ExponentialWithJitter => AddJitter(
                    TimeSpan.FromMilliseconds(policy.BaseDelay.TotalMilliseconds * Math.Pow(policy.BackoffMultiplier, attemptNumber)), 
                    policy.JitterMaxPercent),
                _ => policy.BaseDelay
            };

            // Ensure delay doesn't exceed maximum
            return delay > policy.MaxDelay ? policy.MaxDelay : delay;
        }

        private TimeSpan AddJitter(TimeSpan baseDelay, double jitterMaxPercent)
        {
            if (jitterMaxPercent <= 0) return baseDelay;

            var jitter = _random.NextDouble() * jitterMaxPercent * 2 - jitterMaxPercent; // -jitterMaxPercent to +jitterMaxPercent
            var jitteredDelay = baseDelay.TotalMilliseconds * (1 + jitter);
            return TimeSpan.FromMilliseconds(Math.Max(0, jitteredDelay));
        }
    }

    /// <summary>
    /// Exception thrown when retry attempts are exhausted
    /// </summary>
    public class RetryExhaustedException : Exception
    {
        public RetryExhaustedException(string message, Exception innerException) 
            : base(message, innerException)
        {
        }
    }
}