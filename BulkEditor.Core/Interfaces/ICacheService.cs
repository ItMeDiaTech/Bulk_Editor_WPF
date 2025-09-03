using System;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Core.Interfaces
{
    /// <summary>
    /// Interface for caching service to improve performance
    /// </summary>
    public interface ICacheService
    {
        /// <summary>
        /// Gets a cached value or computes and stores it
        /// </summary>
        Task<T> GetOrSetAsync<T>(string key, Func<Task<T>> factory, TimeSpan? expiry = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// Gets a cached value if it exists
        /// </summary>
        Task<T?> GetAsync<T>(string key, CancellationToken cancellationToken = default);

        /// <summary>
        /// Sets a value in the cache
        /// </summary>
        Task SetAsync<T>(string key, T value, TimeSpan? expiry = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// Removes a value from the cache
        /// </summary>
        Task RemoveAsync(string key, CancellationToken cancellationToken = default);

        /// <summary>
        /// Clears all cached values
        /// </summary>
        Task ClearAsync(CancellationToken cancellationToken = default);

        /// <summary>
        /// Gets cache statistics
        /// </summary>
        CacheStatistics GetStatistics();
    }

    /// <summary>
    /// Cache performance statistics
    /// </summary>
    public class CacheStatistics
    {
        public int TotalEntries { get; set; }
        public int HitCount { get; set; }
        public int MissCount { get; set; }
        public double HitRatio => TotalRequests > 0 ? (double)HitCount / TotalRequests : 0;
        public int TotalRequests => HitCount + MissCount;
        public long MemoryUsageBytes { get; set; }
    }
}