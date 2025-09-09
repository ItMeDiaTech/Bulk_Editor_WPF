using BulkEditor.Core.Interfaces;
using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Simple in-memory cache service implementation for performance optimization
    /// </summary>
    public class MemoryCacheService : ICacheService, IDisposable
    {
        private readonly ConcurrentDictionary<string, CacheEntry> _cache;
        private readonly ILoggingService _logger;
        private readonly CacheStatistics _statistics;
        private readonly object _statsLock = new();
        private readonly Timer _cleanupTimer;

        public MemoryCacheService(ILoggingService logger)
        {
            _cache = new ConcurrentDictionary<string, CacheEntry>();
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _statistics = new CacheStatistics();

            // Cleanup expired entries every 5 minutes
            _cleanupTimer = new Timer(CleanupExpiredEntries, null, TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5));
        }

        public async Task<T> GetOrSetAsync<T>(string key, Func<Task<T>> factory, TimeSpan? expiry = null, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrEmpty(key))
                throw new ArgumentException("Cache key cannot be null or empty", nameof(key));

            // Try to get from cache first
            if (_cache.TryGetValue(key, out var cachedEntry) && !cachedEntry.IsExpired)
            {
                if (cachedEntry.Value is T typedValue)
                {
                    IncrementHitCount();
                    _logger.LogDebug("Cache hit for key: {Key}", key);
                    return typedValue;
                }
            }

            // Cache miss - compute value
            IncrementMissCount();
            _logger.LogDebug("Cache miss for key: {Key}", key);

            var value = await factory();

            // Store in cache with expiry
            var cacheEntry = new CacheEntry
            {
                Value = value,
                ExpiresAt = DateTime.UtcNow.Add(expiry ?? TimeSpan.FromMinutes(30))
            };

            _cache.AddOrUpdate(key, cacheEntry, (k, v) => cacheEntry);
            IncrementTotalEntries();

            return value;
        }

        public Task<T?> GetAsync<T>(string key, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrEmpty(key))
                return Task.FromResult(default(T?));

            if (_cache.TryGetValue(key, out var cachedEntry) && !cachedEntry.IsExpired)
            {
                if (cachedEntry.Value is T typedValue)
                {
                    IncrementHitCount();
                    _logger.LogDebug("Cache hit for key: {Key}", key);
                    return Task.FromResult((T?)typedValue);
                }
            }

            IncrementMissCount();
            return Task.FromResult(default(T?));
        }

        public Task SetAsync<T>(string key, T value, TimeSpan? expiry = null, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrEmpty(key))
                throw new ArgumentException("Cache key cannot be null or empty", nameof(key));

            var cacheEntry = new CacheEntry
            {
                Value = value,
                ExpiresAt = DateTime.UtcNow.Add(expiry ?? TimeSpan.FromMinutes(30))
            };

            _cache.AddOrUpdate(key, cacheEntry, (k, v) => cacheEntry);
            IncrementTotalEntries();

            _logger.LogDebug("Cached value for key: {Key}", key);
            return Task.CompletedTask;
        }

        public Task RemoveAsync(string key, CancellationToken cancellationToken = default)
        {
            if (!string.IsNullOrEmpty(key))
            {
                _cache.TryRemove(key, out _);
                _logger.LogDebug("Removed cache entry for key: {Key}", key);
            }
            return Task.CompletedTask;
        }

        public Task ClearAsync(CancellationToken cancellationToken = default)
        {
            _cache.Clear();

            lock (_statsLock)
            {
                _statistics.TotalEntries = 0;
                _statistics.HitCount = 0;
                _statistics.MissCount = 0;
                _statistics.MemoryUsageBytes = 0;
            }

            _logger.LogInformation("Cache cleared");
            return Task.CompletedTask;
        }

        public CacheStatistics GetStatistics()
        {
            lock (_statsLock)
            {
                // Estimate memory usage (simplified)
                _statistics.MemoryUsageBytes = _cache.Count * 1024; // Rough estimate
                _statistics.TotalEntries = _cache.Count;

                return new CacheStatistics
                {
                    TotalEntries = _statistics.TotalEntries,
                    HitCount = _statistics.HitCount,
                    MissCount = _statistics.MissCount,
                    MemoryUsageBytes = _statistics.MemoryUsageBytes
                };
            }
        }

        private void IncrementHitCount()
        {
            lock (_statsLock)
            {
                _statistics.HitCount++;
            }
        }

        private void IncrementMissCount()
        {
            lock (_statsLock)
            {
                _statistics.MissCount++;
            }
        }

        private void IncrementTotalEntries()
        {
            lock (_statsLock)
            {
                _statistics.TotalEntries = _cache.Count;
            }
        }

        private void CleanupExpiredEntries(object? state)
        {
            try
            {
                var expiredKeys = new List<string>();
                var now = DateTime.UtcNow;

                foreach (var kvp in _cache)
                {
                    if (kvp.Value.IsExpired)
                    {
                        expiredKeys.Add(kvp.Key);
                    }
                }

                foreach (var key in expiredKeys)
                {
                    _cache.TryRemove(key, out _);
                }

                if (expiredKeys.Count > 0)
                {
                    _logger.LogDebug("Cleaned up {Count} expired cache entries", expiredKeys.Count);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during cache cleanup");
            }
        }

        public void Dispose()
        {
            _cleanupTimer?.Dispose();
            _cache?.Clear();
        }

        private class CacheEntry
        {
            public object? Value { get; set; }
            public DateTime ExpiresAt { get; set; }
            public bool IsExpired => DateTime.UtcNow > ExpiresAt;
        }
    }
}