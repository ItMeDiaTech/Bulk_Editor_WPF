using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Core.Interfaces
{
    /// <summary>
    /// Interface for HTTP operations
    /// </summary>
    public interface IHttpService
    {
        /// <summary>
        /// Sends a GET request to the specified URL
        /// </summary>
        Task<HttpResponseMessage> GetAsync(string url, CancellationToken cancellationToken = default);

        /// <summary>
        /// Sends a HEAD request to check URL availability
        /// </summary>
        Task<HttpResponseMessage> HeadAsync(string url, CancellationToken cancellationToken = default);

        /// <summary>
        /// Checks if a URL is accessible
        /// </summary>
        Task<bool> IsUrlAccessibleAsync(string url, CancellationToken cancellationToken = default);

        /// <summary>
        /// Gets the response content as string
        /// </summary>
        Task<string> GetContentAsync(string url, CancellationToken cancellationToken = default);

        /// <summary>
        /// Checks if URL returns a specific status code
        /// </summary>
        Task<bool> CheckUrlStatusAsync(string url, System.Net.HttpStatusCode expectedStatus, CancellationToken cancellationToken = default);

        /// <summary>
        /// Sets timeout for HTTP requests
        /// </summary>
        void SetTimeout(TimeSpan timeout);

        /// <summary>
        /// Sets user agent for HTTP requests
        /// </summary>
        void SetUserAgent(string userAgent);
    }
}