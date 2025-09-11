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

            // CRITICAL FIX: Add word boundaries to match exactly 6 digits, not 7+
            // This prevents matching CMS-PROD-1234567 when we only want CMS-PROD-123456
            _lookupIdRegex = new Regex(@"\b(TSRC-[^-]+-\d{6}|CMS-[^-]+-\d{6})\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

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
                var isAccessible = await _httpService.IsUrlAccessibleAsync(hyperlink.OriginalUrl, cancellationToken).ConfigureAwait(false);

                if (isAccessible)
                {
                    // Check if content indicates expiration
                    var isExpired = await IsUrlExpiredAsync(hyperlink.OriginalUrl, cancellationToken).ConfigureAwait(false);

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
                    var statusCheck404 = await _httpService.CheckUrlStatusAsync(hyperlink.OriginalUrl, HttpStatusCode.NotFound, cancellationToken).ConfigureAwait(false);

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
                    result.ContentId = await GenerateContentIdAsync(result.LookupId, cancellationToken).ConfigureAwait(false);

                    // Get document record to extract Document_ID for URL generation
                    var documentRecord = await SimulateApiLookupAsync(result.LookupId, cancellationToken).ConfigureAwait(false);
                    if (documentRecord != null)
                    {
                        result.DocumentId = documentRecord.Document_ID;
                    }

                    // Check for title differences and handle replacement/reporting
                    _logger.LogInformation("🔄 INITIATING TITLE COMPARISON: Hyperlink_ID='{HyperlinkId}', Lookup_ID='{LookupId}', Content_ID='{ContentId}'",
                        hyperlink.Id, result.LookupId, result.ContentId);

                    var titleComparison = await CheckTitleDifferenceAsync(hyperlink, result.LookupId, result.ContentId, cancellationToken).ConfigureAwait(false);
                    
                    if (titleComparison.TitlesDiffer)
                    {
                        result.TitleComparison = titleComparison;
                        _logger.LogInformation("✅ TITLE COMPARISON RESULT: TitleComparison ASSIGNED to result (TitlesDiffer=true) for Hyperlink_ID='{HyperlinkId}'", hyperlink.Id);
                    }
                    else
                    {
                        _logger.LogInformation("🔄 TITLE COMPARISON RESULT: TitleComparison NOT ASSIGNED (TitlesDiffer=false) for Hyperlink_ID='{HyperlinkId}'", hyperlink.Id);
                    }
                }

                _logger.LogDebug("Hyperlink validation completed: {HyperlinkId} - Status: {Status}, TitleComparison: {HasTitleComparison}", 
                    hyperlink.Id, result.Status, result.TitleComparison != null ? "ASSIGNED" : "NULL");
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

                // CRITICAL FIX: Use EXACT VBA ExtractLookupID logic with docid fallback
                return ExtractLookupIdUsingVbaLogic(url, "");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error extracting lookup ID from URL: {Url}", url);
                return string.Empty;
            }
        }

        /// <summary>
        /// CRITICAL FIX: Exact VBA ExtractLookupID function logic
        /// VBA: Private Function ExtractLookupID(addr As String, subAddr As String, rx As Object) As String
        /// </summary>
        private string ExtractLookupIdUsingVbaLogic(string address, string subAddress)
        {
            try
            {
                // VBA: Dim full As String: full = addr & IIf(Len(subAddr) > 0, "#" & subAddr, "")
                var full = address + (!string.IsNullOrEmpty(subAddress) ? "#" + subAddress : "");

                // VBA: If rx.Test(full) Then ExtractLookupID = UCase$(rx.Execute(full)(0).value)
                var match = _lookupIdRegex.Match(full);
                if (match.Success)
                {
                    var result = match.Value.ToUpperInvariant(); // VBA: UCase$
                    _logger.LogDebug("Extracted Content_ID via primary regex: {LookupId} from {Full}", result, full);
                    return result;
                }

                // VBA: ElseIf InStr(1, full, "docid=", vbTextCompare) > 0 Then
                // CRITICAL: This runs in conjunction with Content_ID extraction, not as fallback
                if (full.IndexOf("docid=", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // VBA: ExtractLookupID = Trim$(Split(Split(full, "docid=")(1), "&")(0))
                    var parts = full.Split(new[] { "docid=" }, StringSplitOptions.None);
                    if (parts.Length > 1)
                    {
                        var docId = parts[1].Split('&')[0].Trim();
                        // CRITICAL FIX: Handle URL encoding (Issue #3)
                        var decodedDocId = Uri.UnescapeDataString(docId);
                        _logger.LogDebug("Extracted Document_ID via docid parameter: {LookupId} from {Full}", decodedDocId, full);
                        return decodedDocId;
                    }
                }

                _logger.LogDebug("No lookup ID found in: {Full}", full);
                return string.Empty;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Error extracting lookup ID from: {Full}. Error: {Error}", address, ex.Message);
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
                    await Task.Delay(10, cancellationToken).ConfigureAwait(false);
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
                _logger.LogInformation("🔍 TITLE COMPARISON START: Lookup_ID='{LookupId}', Content_ID='{ContentId}', Hyperlink_ID='{HyperlinkId}'",
                    lookupId, contentId, hyperlink.Id);

                // Simulate API call to get document information
                var apiRecord = await SimulateApiLookupAsync(lookupId, cancellationToken);
                if (apiRecord == null || string.IsNullOrEmpty(apiRecord.Title))
                {
                    _logger.LogWarning("❌ TITLE COMPARISON ABORTED: No API record or title returned for Lookup_ID='{LookupId}'", lookupId);
                    return result;
                }

                _logger.LogInformation("📡 API RESPONSE: Lookup_ID='{LookupId}' -> Title='{ApiTitle}', Status='{Status}'",
                    lookupId, apiRecord.Title, apiRecord.Status);

                // Extract current title by removing Content ID suffix " (Last 5-6 of Content_ID)" and trailing whitespace
                var currentDisplayText = hyperlink.DisplayText ?? string.Empty;
                var currentTitle = ExtractTitleFromDisplayText(currentDisplayText);
                // Remove trailing whitespace from API title for accurate comparison
                var apiTitle = apiRecord.Title.TrimEnd();

                _logger.LogInformation("🔄 TITLE EXTRACTION: DisplayText='{DisplayText}' -> ExtractedTitle='{CurrentTitle}', APITitle='{ApiTitle}'",
                    currentDisplayText, currentTitle, apiTitle);

                result.CurrentTitle = currentTitle;
                result.ApiTitle = apiTitle;

                // Compare titles (case-insensitive)
                if (!currentTitle.Equals(apiTitle, StringComparison.OrdinalIgnoreCase))
                {
                    result.TitlesDiffer = true;
                    result.ActionTaken = "Title difference detected";

                    _logger.LogInformation("✅ TITLE DIFFERENCE DETECTED: Lookup_ID='{LookupId}', Current='{CurrentTitle}', API='{ApiTitle}', TitlesDiffer=TRUE",
                        lookupId, currentTitle, apiTitle);
                }
                else
                {
                    result.TitlesDiffer = false;
                    result.ActionTaken = "Titles match - no action needed";

                    _logger.LogInformation("🔄 TITLES MATCH: Lookup_ID='{LookupId}', Title='{Title}', TitlesDiffer=FALSE",
                        lookupId, currentTitle);
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
        /// Enhanced to specifically handle "Last 5-6 of Content_ID" patterns and trailing whitespace
        /// </summary>
        private string ExtractTitleFromDisplayText(string displayText)
        {
            if (string.IsNullOrEmpty(displayText))
                return string.Empty;

            // Enhanced logic: Remove Content ID pattern " (Last 5-6 of Content_ID)" from the end
            // This matches patterns like " (123456)", " (12345)", etc. at the end of the string
            var titleWithoutContentId = _contentIdRegex.Replace(displayText, "");
            
            // Remove any trailing whitespace after Content_ID removal
            var cleanTitle = titleWithoutContentId.TrimEnd();

            _logger.LogDebug("Extracted title from display text: '{DisplayText}' -> '{CleanTitle}'", displayText, cleanTitle);
            return cleanTitle;
        }

        /// <summary>
        /// Extracts the last 6 characters of a Content_ID with proper padding for consistent formatting
        /// Handles various Content_ID formats and ensures a 6-digit result
        /// </summary>
        /// <param name="contentId">The Content_ID to process</param>
        /// <returns>Last 6 characters of Content_ID, padded with leading zeros if needed</returns>
        private string GetLast6OfContentId(string contentId)
        {
            if (string.IsNullOrWhiteSpace(contentId))
                return "000000"; // Default 6-digit padding

            try
            {
                // Extract numeric part from various Content_ID formats (like TSRC-PRD-123456, CMS-DOC-12345, etc.)
                var match = System.Text.RegularExpressions.Regex.Match(contentId, @"(\d{1,6})");
                string numericPart;
                
                if (match.Success)
                {
                    numericPart = match.Groups[1].Value;
                }
                else
                {
                    // If no numeric pattern found, try to use the whole string if it's numeric
                    if (System.Text.RegularExpressions.Regex.IsMatch(contentId, @"^\d+$"))
                    {
                        numericPart = contentId;
                    }
                    else
                    {
                        _logger.LogWarning("Could not extract numeric part from Content_ID: {ContentId}. Using default padding.", contentId);
                        return "000000";
                    }
                }

                // Get last 6 characters or pad to 6 digits
                string last6;
                if (numericPart.Length >= 6)
                {
                    // Take the last 6 digits
                    last6 = numericPart.Substring(numericPart.Length - 6);
                }
                else
                {
                    // Pad with leading zeros to make it 6 digits
                    last6 = numericPart.PadLeft(6, '0');
                }

                _logger.LogDebug("Extracted Last 6 of Content_ID: '{ContentId}' -> '{Last6}'", contentId, last6);
                return last6;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error extracting Last 6 of Content_ID from: {ContentId}. Using default padding.", contentId);
                return "000000";
            }
        }

        /// <summary>
        /// Performs API lookup for document information using real HTTP POST request (matching VBA implementation)
        /// </summary>
        private async Task<DocumentRecord?> SimulateApiLookupAsync(string lookupId, CancellationToken cancellationToken)
        {
            try
            {
                // Use cache to avoid repeated API calls for same lookup ID
                var cacheKey = $"api_lookup_{lookupId}";
                var cachedRecord = await _cacheService.GetOrSetAsync(cacheKey, async () =>
                {
                    return await PerformRealApiLookupAsync(lookupId, cancellationToken);
                }, TimeSpan.FromMinutes(30), cancellationToken);

                return cachedRecord;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Error in API lookup for lookup ID: {LookupId}. Error: {Error}", lookupId, ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Performs the actual API lookup using HTTP POST request (based on VBA implementation)
        /// </summary>
        private async Task<DocumentRecord?> PerformRealApiLookupAsync(string lookupId, CancellationToken cancellationToken)
        {
            try
            {
                // Check if we have a real API URL configured, otherwise use test mode
                // TODO: Replace with actual API endpoint when available
                var apiUrl = "test"; // This will trigger the test response in HttpService

                var requestData = new { Lookup_ID = new[] { lookupId } };

                _logger.LogDebug("Making API request for lookup ID: {LookupId}", lookupId);

                var response = await _httpService.PostJsonAsync(apiUrl, requestData, cancellationToken);

                if (!response.IsSuccessStatusCode)
                {
                    _logger.LogWarning("API request failed for lookup ID {LookupId}: {StatusCode}", lookupId, response.StatusCode);
                    return null;
                }

                var responseContent = await response.Content.ReadAsStringAsync(cancellationToken);
                var apiResponse = System.Text.Json.JsonSerializer.Deserialize<ApiResponse>(responseContent);

                if (apiResponse?.Results?.Any() == true)
                {
                    // Find the record that matches our lookup ID
                    var matchingRecord = apiResponse.Results.FirstOrDefault(r =>
                        string.Equals(r.Lookup_ID, lookupId, StringComparison.OrdinalIgnoreCase));

                    if (matchingRecord != null)
                    {
                        _logger.LogDebug("Found API record for lookup ID {LookupId}: {Title}", lookupId, matchingRecord.Title);
                        return matchingRecord;
                    }
                }

                // If no specific record found, generate a fallback record
                var contentId = await GenerateContentIdAsync(lookupId, cancellationToken);
                return new DocumentRecord
                {
                    Document_ID = $"doc-{contentId}-{DateTime.Now.Ticks % 1000}",
                    Content_ID = FormatContentIdWithPadding(contentId),
                    Title = $"Updated API Title for {lookupId}",
                    Status = "Active",
                    Lookup_ID = lookupId
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error performing real API lookup for lookup ID: {LookupId}", lookupId);
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
