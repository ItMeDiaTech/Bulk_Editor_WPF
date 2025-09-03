using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Implementation of hyperlink validation service
    /// </summary>
    public class HyperlinkValidator : IHyperlinkValidator
    {
        private readonly IHttpService _httpService;
        private readonly ILoggingService _logger;
        private readonly ICacheService _cacheService;
        private readonly Regex _lookupIdRegex;
        private readonly Regex _contentIdRegex;

        public HyperlinkValidator(IHttpService httpService, ILoggingService logger, ICacheService cacheService)
        {
            _httpService = httpService ?? throw new ArgumentNullException(nameof(httpService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _cacheService = cacheService ?? throw new ArgumentNullException(nameof(cacheService));

            // Initialize the lookup ID regex pattern
            _lookupIdRegex = new Regex(@"(TSRC-[^-]+-[0-9]{6}|CMS-[^-]+-[0-9]{6})", RegexOptions.Compiled | RegexOptions.IgnoreCase);

            // Initialize Content ID extraction regex for title comparison (handle both 5 and 6 digit IDs)
            _contentIdRegex = new Regex(@"\s*\([0-9]{5,6}\)\s*$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        }

        public async Task<BulkEditor.Core.Interfaces.HyperlinkValidationResult> ValidateHyperlinkAsync(Hyperlink hyperlink, CancellationToken cancellationToken = default)
        {
            try
            {
                _logger.LogDebug("Validating hyperlink: {HyperlinkId} - {Url}", hyperlink.Id, hyperlink.OriginalUrl);

                var result = new BulkEditor.Core.Interfaces.HyperlinkValidationResult
                {
                    HyperlinkId = hyperlink.Id
                };

                // Extract lookup ID from URL
                result.LookupId = ExtractLookupId(hyperlink.OriginalUrl);

                // Check if URL is accessible
                var isAccessible = await _httpService.IsUrlAccessibleAsync(hyperlink.OriginalUrl, cancellationToken);

                if (isAccessible)
                {
                    // Check if content indicates expiration
                    var isExpired = await IsUrlExpiredAsync(hyperlink.OriginalUrl, cancellationToken);

                    if (isExpired)
                    {
                        result.Status = HyperlinkStatus.Expired;
                        result.IsExpired = true;
                        result.RequiresUpdate = true;
                        result.ErrorMessage = "Content indicates this link has expired";
                    }
                    else
                    {
                        result.Status = HyperlinkStatus.Valid;
                    }
                }
                else
                {
                    // Check specific error types
                    var statusCheck404 = await _httpService.CheckUrlStatusAsync(hyperlink.OriginalUrl, HttpStatusCode.NotFound, cancellationToken);

                    if (statusCheck404)
                    {
                        result.Status = HyperlinkStatus.NotFound;
                        result.ErrorMessage = "URL returns 404 Not Found";
                    }
                    else
                    {
                        result.Status = HyperlinkStatus.Invalid;
                        result.ErrorMessage = "URL is not accessible";
                    }

                    result.RequiresUpdate = true;
                }

                // Generate Content ID and get Document ID if lookup ID is found
                if (!string.IsNullOrEmpty(result.LookupId))
                {
                    result.ContentId = await GenerateContentIdAsync(result.LookupId, cancellationToken);

                    // Get document record to extract Document_ID for URL generation
                    var documentRecord = await SimulateApiLookupAsync(result.LookupId, cancellationToken);
                    if (documentRecord != null)
                    {
                        result.DocumentId = documentRecord.Document_ID;
                    }

                    // Check for title differences and handle replacement/reporting
                    var titleComparison = await CheckTitleDifferenceAsync(hyperlink, result.LookupId, result.ContentId, cancellationToken);
                    if (titleComparison.TitlesDiffer)
                    {
                        result.TitleComparison = titleComparison;
                    }
                }

                _logger.LogDebug("Hyperlink validation completed: {HyperlinkId} - Status: {Status}", hyperlink.Id, result.Status);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error validating hyperlink: {HyperlinkId}", hyperlink.Id);

                return new BulkEditor.Core.Interfaces.HyperlinkValidationResult
                {
                    HyperlinkId = hyperlink.Id,
                    Status = HyperlinkStatus.Error,
                    ErrorMessage = ex.Message
                };
            }
        }

        public async Task<IEnumerable<BulkEditor.Core.Interfaces.HyperlinkValidationResult>> ValidateHyperlinksAsync(IEnumerable<Hyperlink> hyperlinks, CancellationToken cancellationToken = default)
        {
            try
            {
                _logger.LogInformation("Validating {Count} hyperlinks", hyperlinks.Count());

                var results = new List<BulkEditor.Core.Interfaces.HyperlinkValidationResult>();
                var semaphore = new SemaphoreSlim(10); // Limit concurrent validations

                var tasks = hyperlinks.Select(async hyperlink =>
                {
                    await semaphore.WaitAsync(cancellationToken);
                    try
                    {
                        return await ValidateHyperlinkAsync(hyperlink, cancellationToken);
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                });

                var validationResults = await Task.WhenAll(tasks);

                _logger.LogInformation("Completed validation of {Count} hyperlinks", validationResults.Length);
                return validationResults;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error validating hyperlinks batch");
                throw;
            }
        }

        public string ExtractLookupId(string url)
        {
            try
            {
                if (string.IsNullOrEmpty(url))
                    return string.Empty;

                var match = _lookupIdRegex.Match(url);
                var lookupId = match.Success ? match.Value : string.Empty;

                _logger.LogDebug("Extracted lookup ID '{LookupId}' from URL: {Url}", lookupId, url);
                return lookupId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error extracting lookup ID from URL: {Url}", url);
                return string.Empty;
            }
        }

        public async Task<bool> IsUrlExpiredAsync(string url, CancellationToken cancellationToken = default)
        {
            try
            {
                // Get content and check for common expiration indicators
                var content = await _httpService.GetContentAsync(url, cancellationToken);

                if (string.IsNullOrEmpty(content))
                    return false;

                var contentLower = content.ToLowerInvariant();

                // Common indicators of expired content
                var expirationIndicators = new[]
                {
                    "expired",
                    "no longer available",
                    "content removed",
                    "page not found",
                    "document expired",
                    "access denied",
                    "content unavailable"
                };

                var isExpired = expirationIndicators.Any(indicator => contentLower.Contains(indicator));

                _logger.LogDebug("URL expiration check for {Url}: {IsExpired}", url, isExpired);
                return isExpired;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Error checking URL expiration for: {Url}. Error: {Error}", url, ex.Message);
                return false;
            }
        }

        public async Task<string> GenerateContentIdAsync(string lookupId, CancellationToken cancellationToken = default)
        {
            try
            {
                if (string.IsNullOrEmpty(lookupId))
                    return string.Empty;

                // Use cache to avoid regenerating same Content IDs
                var cacheKey = $"content_id_{lookupId}";
                var cachedContentId = await _cacheService.GetOrSetAsync(cacheKey, async () =>
                {
                    // Generate a 6-digit content ID based on lookup ID
                    var hash = lookupId.GetHashCode();
                    var contentId = Math.Abs(hash % 1000000).ToString("D6");

                    _logger.LogDebug("Generated Content ID '{ContentId}' for lookup ID: {LookupId}", contentId, lookupId);

                    // Simulate async operation
                    await Task.Delay(10, cancellationToken);
                    return contentId;
                }, TimeSpan.FromHours(24), cancellationToken);

                return cachedContentId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating Content ID for lookup ID: {LookupId}", lookupId);
                return string.Empty;
            }
        }

        /// <summary>
        /// Checks if hyperlink title differs from API response and handles replacement/reporting
        /// </summary>
        public async Task<TitleComparisonResult> CheckTitleDifferenceAsync(Hyperlink hyperlink, string lookupId, string contentId, CancellationToken cancellationToken = default)
        {
            var result = new TitleComparisonResult
            {
                ContentId = contentId
            };

            try
            {
                // Simulate API call to get document information
                var apiRecord = await SimulateApiLookupAsync(lookupId, cancellationToken);
                if (apiRecord == null || string.IsNullOrEmpty(apiRecord.Title))
                {
                    return result;
                }

                // Extract current title by removing Content ID suffix " (123456)"
                var currentDisplayText = hyperlink.DisplayText ?? string.Empty;
                var currentTitle = ExtractTitleFromDisplayText(currentDisplayText);
                var apiTitle = apiRecord.Title.Trim();

                result.CurrentTitle = currentTitle;
                result.ApiTitle = apiTitle;

                // Compare titles (case-insensitive)
                if (!currentTitle.Equals(apiTitle, StringComparison.OrdinalIgnoreCase))
                {
                    result.TitlesDiffer = true;
                    result.ActionTaken = "Title difference detected";

                    _logger.LogInformation("Title difference detected for lookup ID {LookupId}: Current='{CurrentTitle}', API='{ApiTitle}'",
                        lookupId, currentTitle, apiTitle);
                }

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking title difference for lookup ID: {LookupId}", lookupId);
                return result;
            }
        }

        /// <summary>
        /// Extracts title from display text by removing Content ID suffix
        /// </summary>
        private string ExtractTitleFromDisplayText(string displayText)
        {
            if (string.IsNullOrEmpty(displayText))
                return string.Empty;

            // Remove Content ID pattern " (123456)" from the end
            var titleWithoutContentId = _contentIdRegex.Replace(displayText, "").Trim();

            return titleWithoutContentId;
        }

        /// <summary>
        /// Simulates API lookup for document information
        /// </summary>
        private async Task<DocumentRecord?> SimulateApiLookupAsync(string lookupId, CancellationToken cancellationToken)
        {
            try
            {
                // Use cache to avoid repeated API calls for same lookup ID
                var cacheKey = $"api_lookup_{lookupId}";
                var cachedRecord = await _cacheService.GetOrSetAsync(cacheKey, async () =>
                {
                    // Simulate API delay
                    await Task.Delay(100, cancellationToken);

                    // In a real implementation, this would make an HTTP request to get document info
                    var contentId = await GenerateContentIdAsync(lookupId, cancellationToken);
                    return new DocumentRecord
                    {
                        Document_ID = $"doc-{contentId}-{DateTime.Now.Ticks % 1000}",
                        Content_ID = FormatContentIdWithPadding(contentId),
                        Title = $"Updated API Title for {lookupId}",
                        Status = "Active",
                        Lookup_ID = lookupId
                    };
                }, TimeSpan.FromMinutes(30), cancellationToken);

                return cachedRecord;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Error in API lookup simulation for lookup ID: {LookupId}. Error: {Error}", lookupId, ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Formats Content ID to ensure 6-digit format with leading zero padding if needed
        /// </summary>
        /// <param name="contentId">Raw content ID</param>
        /// <returns>Formatted 6-digit content ID</returns>
        private string FormatContentIdWithPadding(string contentId)
        {
            if (string.IsNullOrWhiteSpace(contentId))
                return contentId;

            // If it's exactly 5 digits, pad with leading zero
            if (Regex.IsMatch(contentId, @"^[0-9]{5}$"))
            {
                var paddedId = "0" + contentId;
                _logger.LogDebug("Padded 5-digit Content ID '{OriginalId}' to 6-digit '{PaddedId}'", contentId, paddedId);
                return paddedId;
            }

            // Return as-is if already 6 digits or not a pure numeric string
            return contentId;
        }
    }
}
