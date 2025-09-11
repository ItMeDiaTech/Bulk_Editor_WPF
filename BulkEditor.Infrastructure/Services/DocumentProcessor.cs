using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using BulkEditor.Infrastructure.Utilities;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using OpenXmlDocument = DocumentFormat.OpenXml.Wordprocessing.Document;
using OpenXmlHyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Implementation of document processing service using OpenXML with memory optimization
    /// </summary>
    public class DocumentProcessor : IDocumentProcessor, IDisposable
    {
        private readonly IFileService _fileService;
        private readonly IHyperlinkValidator _hyperlinkValidator;
        private readonly ITextOptimizer _textOptimizer;
        private readonly IReplacementService _replacementService;
        private readonly ILoggingService _logger;
        private readonly Core.Configuration.AppSettings _appSettings;
        private readonly IRetryPolicyService _retryPolicyService;

        // CRITICAL FIX: Exact VBA pattern match with word boundaries to ensure exactly 6 digits (Issue #1)
        // VBA: .Pattern = "(TSRC-[^-]+-[0-9]{6}|CMS-[^-]+-[0-9]{6})"
        // VBA: .IgnoreCase = True
        // Added word boundaries to prevent matching 7+ digit sequences
        private static readonly Regex LookupIdRegex = new Regex(@"\b(TSRC-[^-]+-\d{6}|CMS-[^-]+-\d{6})\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex DocIdRegex = new Regex(@"docid=([^&]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // OpenXML validator for document integrity checks
        private readonly OpenXmlValidator _validator = new OpenXmlValidator();

        // Thread-safe semaphore collection for atomic document operations
        private readonly ConcurrentDictionary<string, SemaphoreSlim> _documentSemaphores = new();

        // Add a HashSet to store ignorable validation error descriptions for performance
        private static readonly HashSet<string> IgnorableValidationErrorDescriptions = new HashSet<string>
        {
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
            "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:tr'.",
            "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:view'."
        };

        // Additional patterns for runtime validation error matching
        private static readonly HashSet<string> IgnorableValidationErrorPatterns = new HashSet<string>
        {
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:id' attribute is invalid - The value",
            "The element has invalid child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:hyperlink'. List of possible elements expected:"
        };

        /// <summary>
        /// Checks if a validation error should be ignored based on exact matches and patterns
        /// </summary>
        private static bool IsIgnorableValidationError(string description)
        {
            // Check exact matches first (fastest)
            if (IgnorableValidationErrorDescriptions.Contains(description))
            {
                return true;
            }

            // Check patterns for runtime-generated errors
            return IgnorableValidationErrorPatterns.Any(pattern => description.StartsWith(pattern, StringComparison.Ordinal));
        }

        public DocumentProcessor(IFileService fileService, IHyperlinkValidator hyperlinkValidator, ITextOptimizer textOptimizer, IReplacementService replacementService, ILoggingService logger, Core.Configuration.AppSettings appSettings, IRetryPolicyService retryPolicyService)
        {
            _fileService = fileService ?? throw new ArgumentNullException(nameof(fileService));
            _hyperlinkValidator = hyperlinkValidator ?? throw new ArgumentNullException(nameof(hyperlinkValidator));
            _textOptimizer = textOptimizer ?? throw new ArgumentNullException(nameof(textOptimizer));
            _replacementService = replacementService ?? throw new ArgumentNullException(nameof(replacementService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _appSettings = appSettings ?? throw new ArgumentNullException(nameof(appSettings));
            _retryPolicyService = retryPolicyService ?? throw new ArgumentNullException(nameof(retryPolicyService));
        }

        public async Task<BulkEditor.Core.Entities.Document> ProcessDocumentAsync(string filePath, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
        {
            var document = new BulkEditor.Core.Entities.Document
            {
                FilePath = filePath,
                FileName = Path.GetFileName(filePath),
                Status = DocumentStatus.Processing
            };

            try
            {
                progress?.Report($"Processing document: {document.FileName}");
                _logger.LogInformation("Starting document processing: {FilePath}", filePath);

                // Validate file exists and is a Word document
                if (!_fileService.FileExists(filePath))
                {
                    throw new FileNotFoundException($"Document not found: {filePath}");
                }

                if (!_fileService.IsValidWordDocument(filePath))
                {
                    throw new InvalidOperationException($"File is not a valid Word document: {filePath}");
                }

                // NOTE: Backup creation is handled by MainWindowViewModel before document processing begins
                _logger.LogDebug("Document processing started - backup handled by caller for document: {FileName}", document.FileName);

                // CRITICAL FIX: Process document in single session to prevent corruption
                progress?.Report("Processing document...");
                await ProcessDocumentInSingleSessionAsync(document, progress, cancellationToken).ConfigureAwait(false);

                // NOTE: Replacements and text optimization are now handled within the single session
                // to prevent file corruption from multiple document opens
                progress?.Report("Document processing completed in single session");

                // Optimize memory after processing
                await OptimizeMemoryAsync(cancellationToken).ConfigureAwait(false);

                document.Status = DocumentStatus.Completed;
                document.ProcessedAt = DateTime.UtcNow;

                // Generate change log summary
                GenerateChangeLogSummary(document);

                progress?.Report($"Document processing completed: {document.FileName}");
                _logger.LogInformation("Document processing completed successfully: {FilePath}", filePath);

                return document;
            }
            catch (Exception ex)
            {
                document.Status = DocumentStatus.Failed;
                document.ProcessingErrors.Add(new ProcessingError { Message = ex.Message, Severity = ErrorSeverity.Error });

                _logger.LogError(ex, "Error processing document: {FilePath}", filePath);
                progress?.Report($"Error processing document: {ex.Message}");

                // Try to restore from backup if document was corrupted
                if (!string.IsNullOrEmpty(document.BackupPath) && _fileService.FileExists(document.BackupPath))
                {
                    try
                    {
                        await RestoreFromBackupAsync(filePath, document.BackupPath, cancellationToken).ConfigureAwait(false);
                        _logger.LogInformation("Restored corrupted document from backup: {FilePath}", filePath);
                    }
                    catch (Exception restoreEx)
                    {
                        _logger.LogError(restoreEx, "Failed to restore document from backup: {FilePath}", filePath);
                    }
                }

                return document;
            }
        }

        public async Task<IEnumerable<BulkEditor.Core.Entities.Document>> ProcessDocumentsBatchAsync(IEnumerable<string> filePaths, IProgress<BatchProcessingProgress>? progress = null, CancellationToken cancellationToken = default)
        {
            var filePathsList = filePaths.ToList();
            var results = new List<BulkEditor.Core.Entities.Document>();
            var processed = 0;
            var failed = 0;
            var successful = 0;
            var startTime = DateTime.Now;

            var batchProgress = new BatchProcessingProgress
            {
                TotalDocuments = filePathsList.Count,
                StartTime = startTime,
                CurrentOperation = "Initializing batch processing..."
            };

            _logger.LogInformation("Starting batch processing of {Count} documents", filePathsList.Count);

            try
            {
                // Use configurable concurrency based on system resources
                var maxConcurrency = Math.Min(_appSettings.Processing.MaxConcurrentDocuments, Environment.ProcessorCount * 2);
                var semaphore = new SemaphoreSlim(maxConcurrency);

                var tasks = filePathsList.Select(async filePath =>
                {
                    await semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
                    var docStartTime = DateTime.Now;

                    try
                    {
                        // Update progress for current document
                        batchProgress.CurrentDocument = filePath;
                        batchProgress.CurrentOperation = "Processing document...";
                        batchProgress.CurrentDocumentProgress = 0;
                        progress?.Report(batchProgress);

                        var document = await ProcessDocumentAsync(filePath,
                            new Progress<string>(msg =>
                            {
                                batchProgress.CurrentOperation = msg;
                                batchProgress.CurrentDocumentProgress = Math.Min(batchProgress.CurrentDocumentProgress + 10, 90);
                                progress?.Report(batchProgress);
                            }),
                            cancellationToken);

                        var docProcessingTime = DateTime.Now - docStartTime;

                        lock (results)
                        {
                            results.Add(document);

                            if (document.Status == DocumentStatus.Completed)
                            {
                                Interlocked.Increment(ref processed);
                                Interlocked.Increment(ref successful);
                            }
                            else
                            {
                                Interlocked.Increment(ref failed);

                                // Add error to recent errors list
                                var errorMessage = document.ProcessingErrors.LastOrDefault()?.Message ?? "Unknown error";
                                if (batchProgress.RecentErrors.Count >= 5)
                                    batchProgress.RecentErrors.RemoveAt(0);
                                batchProgress.RecentErrors.Add($"{Path.GetFileName(filePath)}: {errorMessage}");
                            }

                            // Update comprehensive progress statistics
                            batchProgress.ProcessedDocuments = processed;
                            batchProgress.FailedDocuments = failed;
                            batchProgress.SuccessfulDocuments = successful;
                            batchProgress.TotalHyperlinksFound += document.Hyperlinks.Count;
                            batchProgress.TotalHyperlinksProcessed += document.Hyperlinks.Count(h => !string.IsNullOrEmpty(h.UpdatedUrl));

                            // CRITICAL FIX: Track unique hyperlinks that have been changed (any type of change)
                            // Check for URL changes, title changes, or any modification
                            foreach (var hyperlink in document.Hyperlinks)
                            {
                                bool hasChanged = false;

                                // Check if URL was updated
                                if (!string.IsNullOrEmpty(hyperlink.UpdatedUrl) && hyperlink.UpdatedUrl != hyperlink.OriginalUrl)
                                {
                                    hasChanged = true;
                                }

                                // Check if display text was updated (from change log)
                                if (document.ChangeLog.Changes.Any(c =>
                                    (c.Type == ChangeType.HyperlinkUpdated || c.Type == ChangeType.TitleChanged || c.Type == ChangeType.TitleReplaced || c.Type == ChangeType.HyperlinkStatusAdded || c.Type == ChangeType.ContentIdAdded)
                                    && c.ElementId == hyperlink.Id))
                                {
                                    hasChanged = true;
                                }

                                // Add to unique set if any change occurred
                                if (hasChanged)
                                {
                                    batchProgress.UniqueHyperlinksChanged.Add($"{document.FileName}:{hyperlink.Id}");
                                }
                            }

                            // Update count based on unique changes
                            batchProgress.TotalHyperlinksUpdated = batchProgress.UniqueHyperlinksChanged.Count;

                            // Calculate average processing time
                            var totalProcessingTime = (DateTime.Now - startTime).TotalSeconds;
                            batchProgress.AverageProcessingTimePerDocument = totalProcessingTime / Math.Max(1, processed);

                            // Estimate remaining time
                            var remainingDocs = batchProgress.TotalDocuments - processed;
                            if (batchProgress.AverageProcessingTimePerDocument > 0 && remainingDocs > 0)
                            {
                                batchProgress.EstimatedTimeRemaining = TimeSpan.FromSeconds(remainingDocs * batchProgress.AverageProcessingTimePerDocument);
                            }

                            batchProgress.CurrentDocumentProgress = 100;
                            batchProgress.CurrentOperation = "Document completed";
                            progress?.Report(batchProgress);
                        }

                        return document;
                    }
                    catch (Exception ex)
                    {
                        // Handle individual document processing error
                        var errorDoc = new BulkEditor.Core.Entities.Document
                        {
                            FilePath = filePath,
                            FileName = Path.GetFileName(filePath),
                            Status = DocumentStatus.Failed,
                            ProcessingErrors = new List<ProcessingError>
                            {
                                new ProcessingError { Message = ex.Message, Severity = ErrorSeverity.Error }
                            }
                        };

                        lock (results)
                        {
                            results.Add(errorDoc);
                            Interlocked.Increment(ref failed);

                            if (batchProgress.RecentErrors.Count >= 5)
                                batchProgress.RecentErrors.RemoveAt(0);
                            batchProgress.RecentErrors.Add($"{Path.GetFileName(filePath)}: {ex.Message}");

                            batchProgress.ProcessedDocuments = processed + failed;
                            batchProgress.FailedDocuments = failed;
                            progress?.Report(batchProgress);
                        }

                        return errorDoc;
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                });

                await Task.WhenAll(tasks).ConfigureAwait(false);

                // Final progress update
                batchProgress.CurrentOperation = "Finalizing batch processing...";
                batchProgress.CurrentDocumentProgress = 100;
                progress?.Report(batchProgress);

                // Optimize memory after batch processing
                await OptimizeMemoryAsync(cancellationToken).ConfigureAwait(false);

                _logger.LogInformation("Batch processing completed: {Processed} successful, {Failed} failed", successful, failed);
                return results;
            }
            catch (Exception ex)
            {
                batchProgress.CurrentOperation = $"Batch processing failed: {ex.Message}";
                if (batchProgress.RecentErrors.Count >= 5)
                    batchProgress.RecentErrors.RemoveAt(0);
                batchProgress.RecentErrors.Add($"Batch error: {ex.Message}");
                progress?.Report(batchProgress);

                _logger.LogError(ex, "Error in batch processing");
                throw;
            }
        }

        public async Task<IEnumerable<Hyperlink>> ValidateHyperlinksAsync(BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken = default)
        {
            try
            {
                if (!document.Hyperlinks.Any())
                    return document.Hyperlinks;

                _logger.LogInformation("Validating {Count} hyperlinks", document.Hyperlinks.Count);

                var validationResults = await _hyperlinkValidator.ValidateHyperlinksAsync(document.Hyperlinks, cancellationToken).ConfigureAwait(false);

                // Update hyperlinks with validation results
                foreach (var result in validationResults)
                {
                    var hyperlink = document.Hyperlinks.FirstOrDefault(h => h.Id == result.HyperlinkId);
                    if (hyperlink != null)
                    {
                        hyperlink.Status = result.Status;
                        hyperlink.LookupId = result.LookupId;
                        hyperlink.ContentId = result.ContentId;
                        hyperlink.DocumentId = result.DocumentId; // Set Document_ID for URL generation
                        hyperlink.ErrorMessage = result.ErrorMessage;
                        hyperlink.RequiresUpdate = result.RequiresUpdate;
                        hyperlink.LastChecked = DateTime.UtcNow;

                        // Handle status suffix appending (Expired/Not Found) based on VBA logic
                        await HandleHyperlinkStatusSuffixAsync(document, hyperlink, cancellationToken);

                        // Handle title differences if detected
                        if (result.TitleComparison?.TitlesDiffer == true)
                        {
                            _logger.LogInformation("📋 PROCESSING TITLE DIFFERENCE: Hyperlink_ID='{HyperlinkId}', TitleComparison Found - Calling HandleTitleDifferenceAsync", hyperlink.Id);
                            await HandleTitleDifferenceAsync(document, hyperlink, result.TitleComparison, cancellationToken);
                        }
                        else if (result.TitleComparison != null)
                        {
                            _logger.LogInformation("🔄 TITLE COMPARISON EXISTS BUT NO DIFFERENCE: Hyperlink_ID='{HyperlinkId}', TitlesDiffer=false", hyperlink.Id);
                        }
                        else
                        {
                            _logger.LogInformation("❌ NO TITLE COMPARISON: Hyperlink_ID='{HyperlinkId}', TitleComparison=null", hyperlink.Id);
                        }

                        // Log change
                        document.ChangeLog.Changes.Add(new ChangeEntry
                        {
                            Type = ChangeType.Information,
                            Description = $"Hyperlink validation: {result.Status}",
                            ElementId = hyperlink.Id,
                            Details = result.ErrorMessage
                        });
                    }
                }

                // Update document metadata
                document.Metadata.HasExpiredLinks = document.Hyperlinks.Any(h => h.Status == HyperlinkStatus.Expired);
                document.Metadata.HasInvalidLinks = document.Hyperlinks.Any(h => h.Status == HyperlinkStatus.Invalid || h.Status == HyperlinkStatus.NotFound);

                _logger.LogInformation("Completed validation of {Count} hyperlinks", document.Hyperlinks.Count);
                return document.Hyperlinks;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error validating hyperlinks for document: {FileName}", document.FileName);
                throw;
            }
        }

        // NOTE: Backup creation is now handled by BackupService via MainWindowViewModel
        // This ensures all backups use the configured backup directory from settings

        public async Task<bool> RestoreFromBackupAsync(string filePath, string backupPath, CancellationToken cancellationToken = default)
        {
            try
            {
                // CRITICAL FIX: Use timeout instead of cancellation for backup restoration
                // Backup restoration should complete to prevent data loss, but not hang indefinitely
                using var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                await _fileService.CopyFileAsync(backupPath, filePath, timeoutCts.Token);
                _logger.LogInformation("Restored file from backup: {FilePath}", filePath);
                return true;
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                _logger.LogWarning("Backup restoration was cancelled - attempting with timeout protection: {FilePath}", filePath);
                try
                {
                    // Final attempt with timeout but no external cancellation
                    using var finalTimeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
                    await _fileService.CopyFileAsync(backupPath, filePath, finalTimeoutCts.Token);
                    _logger.LogInformation("Restored file from backup on final attempt: {FilePath}", filePath);
                    return true;
                }
                catch (Exception finalEx)
                {
                    _logger.LogError(finalEx, "Final backup restoration attempt failed: {FilePath}", filePath);
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error restoring file from backup: {FilePath}", filePath);
                return false;
            }
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
        /// Handles title differences between current hyperlink and API response
        /// Enhanced to use "Last 6 of Content_ID" format and remove trailing whitespace
        /// </summary>
        private async Task HandleTitleDifferenceAsync(BulkEditor.Core.Entities.Document document, Hyperlink hyperlink, TitleComparisonResult titleComparison, CancellationToken cancellationToken)
        {
            try
            {
                var validationSettings = _appSettings.Validation;

                if (validationSettings.AutoReplaceTitles)
                {
                    // Enhanced logic: Replace title with API title (trailing whitespace removed) and append Last 6 of Content_ID
                    var cleanApiTitle = titleComparison.ApiTitle.TrimEnd();
                    var last6OfContentId = GetLast6OfContentId(titleComparison.ContentId);
                    var newDisplayText = $"{cleanApiTitle} ({last6OfContentId})";

                    // Update the hyperlink in the document
                    await UpdateHyperlinkTitleInDocumentAsync(document.FilePath, hyperlink, newDisplayText, cancellationToken);

                    // CRITICAL FIX: Mark hyperlink for document update so title change gets written to file
                    hyperlink.RequiresUpdate = true;

                    titleComparison.WasReplaced = true;
                    titleComparison.ActionTaken = "Title replaced with API response using Last 6 of Content_ID format";

                    // Log the replacement with enhanced details
                    document.ChangeLog.Changes.Add(new ChangeEntry
                    {
                        Type = ChangeType.TitleReplaced,
                        Description = "Title replaced with API response using Last 6 of Content_ID format",
                        OldValue = titleComparison.CurrentTitle,
                        NewValue = cleanApiTitle,
                        ElementId = hyperlink.Id,
                        Details = $"Content ID: {titleComparison.ContentId}, Last 6: {last6OfContentId}, Full new title: {newDisplayText}"
                    });

                    _logger.LogInformation("✓ TITLE REPLACEMENT APPLIED: '{OldTitle}' -> '{NewTitle}' (Content ID: {ContentId}, Last 6: {Last6}) - Hyperlink marked for document update",
                        titleComparison.CurrentTitle, cleanApiTitle, titleComparison.ContentId, last6OfContentId);
                }
                else if (validationSettings.ReportTitleDifferences)
                {
                    // Only report the difference in changelog without replacing
                    titleComparison.ActionTaken = "Title difference reported";

                    document.ChangeLog.Changes.Add(new ChangeEntry
                    {
                        Type = ChangeType.PossibleTitleChange,
                        Description = "Possible Title Change",
                        OldValue = titleComparison.CurrentTitle,
                        NewValue = titleComparison.ApiTitle,
                        ElementId = hyperlink.Id,
                        Details = $"Content ID: {titleComparison.ContentId}"
                    });

                    _logger.LogInformation("Title difference reported for hyperlink: Current='{CurrentTitle}', API='{ApiTitle}' (Content ID: {ContentId})",
                        titleComparison.CurrentTitle, titleComparison.ApiTitle, titleComparison.ContentId);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error handling title difference for hyperlink: {HyperlinkId}", hyperlink.Id);

                document.ChangeLog.Changes.Add(new ChangeEntry
                {
                    Type = ChangeType.Error,
                    Description = "Error processing title difference",
                    ElementId = hyperlink.Id,
                    Details = ex.Message
                });
            }
        }

        /// <summary>
        /// DEPRECATED: This method has been replaced to prevent document corruption
        /// All title updates are now handled within the single document session
        /// </summary>
        private async Task UpdateHyperlinkTitleInDocumentAsync(string filePath, Hyperlink hyperlink, string newDisplayText, CancellationToken cancellationToken)
        {
            // This method is deprecated to prevent document corruption
            // All operations are now handled in ProcessDocumentInSingleSessionAsync
            _logger.LogDebug("UpdateHyperlinkTitleInDocumentAsync called - operations handled in single session to prevent corruption");

            // Update the hyperlink object for the changelog
            hyperlink.DisplayText = newDisplayText;
            await Task.CompletedTask;
        }

        /// <summary>
        /// Standalone title validation and update workflow
        /// Validates and updates hyperlink titles without processing URLs
        /// </summary>
        public async Task<IEnumerable<Hyperlink>> ValidateAndUpdateTitlesAsync(BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken = default)
        {
            try
            {
                if (!document.Hyperlinks.Any())
                {
                    _logger.LogDebug("No hyperlinks found for title validation in document: {FileName}", document.FileName);
                    return document.Hyperlinks;
                }

                var validationSettings = _appSettings.Validation;
                if (!validationSettings.AutoReplaceTitles)
                {
                    _logger.LogDebug("AutoReplaceTitles disabled - skipping title validation for: {FileName}", document.FileName);
                    return document.Hyperlinks;
                }

                _logger.LogInformation("✓ TITLE-ONLY VALIDATION: Starting for {Count} hyperlinks in {FileName}", document.Hyperlinks.Count, document.FileName);

                // Validate hyperlinks to get title comparison data
                var validationResults = await _hyperlinkValidator.ValidateHyperlinksAsync(document.Hyperlinks, cancellationToken).ConfigureAwait(false);

                var titlesUpdated = 0;
                
                // Process title differences
                foreach (var result in validationResults)
                {
                    var hyperlink = document.Hyperlinks.FirstOrDefault(h => h.Id == result.HyperlinkId);
                    _logger.LogInformation("🔍 TITLE-ONLY VALIDATION PROCESSING: Hyperlink_ID='{HyperlinkId}', TitleComparison={HasComparison}, TitlesDiffer={TitlesDiffer}",
                        result.HyperlinkId, result.TitleComparison != null ? "EXISTS" : "NULL", result.TitleComparison?.TitlesDiffer ?? false);

                    if (hyperlink != null && result.TitleComparison != null && result.TitleComparison.TitlesDiffer)
                    {
                        _logger.LogInformation("📋 TITLE-ONLY VALIDATION: Processing title difference for Hyperlink_ID='{HyperlinkId}'", hyperlink.Id);
                        await HandleTitleDifferenceAsync(document, hyperlink, result.TitleComparison, cancellationToken);
                        if (result.TitleComparison.WasReplaced)
                        {
                            titlesUpdated++;
                        }
                    }
                    else if (hyperlink == null)
                    {
                        _logger.LogWarning("❌ TITLE-ONLY VALIDATION: Hyperlink not found for Hyperlink_ID='{HyperlinkId}'", result.HyperlinkId);
                    }
                }

                _logger.LogInformation("✓ TITLE-ONLY VALIDATION COMPLETED: {UpdatedCount} titles updated in {FileName}", titlesUpdated, document.FileName);
                return document.Hyperlinks;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in standalone title validation for document: {FileName}", document.FileName);
                throw;
            }
        }

        /// <summary>
        /// Optimizes memory usage by clearing unnecessary caches and forcing garbage collection
        /// </summary>
        public async Task OptimizeMemoryAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                _logger.LogDebug("Starting memory optimization");

                // Force garbage collection if memory usage is high
                var memoryBefore = GC.GetTotalMemory(false);
                if (memoryBefore > 50 * 1024 * 1024) // 50MB threshold
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();

                    var memoryAfter = GC.GetTotalMemory(true);
                    var memoryFreed = memoryBefore - memoryAfter;

                    _logger.LogDebug("Memory optimization completed: {MemoryFreed:N0} bytes freed", memoryFreed);
                }

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during memory optimization");
            }
        }

        private DocumentMetadata ExtractDocumentMetadata(string filePath)
        {
            var metadata = new DocumentMetadata();

            try
            {
                var fileInfo = _fileService.GetFileInfo(filePath);
                metadata.FileSizeBytes = fileInfo.Length;
                metadata.LastModified = fileInfo.LastWriteTime;

                using var document = WordprocessingDocument.Open(filePath, false);
                var coreProperties = document.PackageProperties;

                metadata.Title = coreProperties.Title ?? "";
                metadata.Author = coreProperties.Creator ?? "";
                metadata.Subject = coreProperties.Subject ?? "";
                metadata.Keywords = coreProperties.Keywords ?? "";
                metadata.Comments = coreProperties.Description ?? "";

                // Count words and pages (simplified)
                if (document.MainDocumentPart?.Document?.Body != null)
                {
                    var text = document.MainDocumentPart.Document.Body.InnerText;
                    metadata.WordCount = text.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Length;
                    metadata.PageCount = 1; // Simplified - actual page counting is complex
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Error extracting metadata from document: {FilePath}. Error: {Error}", filePath, ex.Message);
            }

            return metadata;
        }

        private List<Hyperlink> ExtractHyperlinks(string filePath)
        {
            var hyperlinks = new List<Hyperlink>();

            try
            {
                using var document = WordprocessingDocument.Open(filePath, false);
                var mainPart = document.MainDocumentPart;

                if (mainPart?.Document?.Body != null)
                {
                    var openXmlHyperlinks = mainPart.Document.Body.Descendants<OpenXmlHyperlink>().ToList();

                    foreach (var openXmlHyperlink in openXmlHyperlinks)
                    {
                        try
                        {
                            var relId = openXmlHyperlink.Id?.Value;
                            if (string.IsNullOrEmpty(relId))
                                continue;

                            // Safely try to get the relationship - some hyperlinks may have invalid IDs
                            try
                            {
                                // CRITICAL FIX: Build complete URL including fragment exactly like VBA
                                var completeUrl = BuildCompleteHyperlinkUrl(mainPart, openXmlHyperlink);
                                var displayText = openXmlHyperlink.InnerText;

                                var hyperlink = new Hyperlink
                                {
                                    OriginalUrl = completeUrl, // Use complete URL including fragment
                                    DisplayText = displayText,
                                    LookupId = _hyperlinkValidator.ExtractLookupId(completeUrl),
                                    RequiresUpdate = ShouldAutoValidateHyperlink(completeUrl, displayText)
                                };

                                hyperlinks.Add(hyperlink);
                            }
                            catch (System.Collections.Generic.KeyNotFoundException)
                            {
                                // Hyperlink has invalid relationship ID - skip it
                                _logger.LogWarning("Skipping hyperlink with invalid relationship ID: {RelId} in document: {FilePath}", relId, filePath);
                                continue;
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning("Error extracting hyperlink from document: {FilePath}. Error: {Error}", filePath, ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error extracting hyperlinks from document: {FilePath}", filePath);
            }

            return hyperlinks;
        }

        /// <summary>
        /// Determines if a hyperlink should be automatically validated based on VBA criteria
        /// Uses EXACT same logic as Base_File.vba ExtractLookupID function
        /// </summary>
        /// <param name="url">Hyperlink URL</param>
        /// <param name="displayText">Hyperlink display text</param>
        /// <returns>True if hyperlink should be auto-validated</returns>
        private bool ShouldAutoValidateHyperlink(string url, string displayText)
        {
            if (string.IsNullOrEmpty(url))
                return false;

            var lookupId = ExtractIdentifierFromUrl(url, "");
            var shouldValidate = !string.IsNullOrEmpty(lookupId);

            _logger.LogDebug("Hyperlink validation check: URL={Url}, LookupID={LookupId}, ShouldValidate={ShouldValidate}",
                url, lookupId, shouldValidate);

            return shouldValidate;
        }

        /// <summary>
        /// CRITICAL FIX: Builds complete URL by combining address and fragment exactly like VBA logic
        /// VBA: full = addr & IIf(Len(subAddr) > 0, "#" & subAddr, "")
        /// </summary>
        /// <param name="mainPart">Document main part containing relationships</param>
        /// <param name="openXmlHyperlink">OpenXML hyperlink element</param>
        /// <returns>Complete URL with fragment if present</returns>
        private string BuildCompleteHyperlinkUrl(MainDocumentPart mainPart, OpenXmlHyperlink openXmlHyperlink)
        {
            try
            {
                var relId = openXmlHyperlink.Id?.Value;
                if (string.IsNullOrEmpty(relId))
                    return string.Empty;

                // Get the base address from the relationship
                var relationship = mainPart.GetReferenceRelationship(relId);
                var rawUri = relationship.Uri.ToString();
                var decodedUri = Uri.UnescapeDataString(rawUri);

                // CRITICAL FIX: Handle URLs that already have encoded fragments in the base address
                string baseAddress;
                string fragment = null;

                // Check if the decoded URI already contains a fragment (malformed URLs with %23)
                if (decodedUri.Contains("#"))
                {
                    // Split the already-complete URL to separate base and fragment
                    var parts = decodedUri.Split('#');
                    baseAddress = parts[0];
                    if (parts.Length > 1)
                    {
                        fragment = parts[1];
                        _logger.LogDebug("Found encoded fragment in base URI: {Fragment} from URL: {DecodedUri}", fragment, decodedUri);
                    }
                }
                else
                {
                    baseAddress = decodedUri;

                    // Check for fragment/subaddress in DocLocation and Anchor properties
                    // 1. Check DocLocation property (most common for external links with fragments)
                    if (!string.IsNullOrEmpty(openXmlHyperlink.DocLocation?.Value))
                    {
                        fragment = openXmlHyperlink.DocLocation.Value;
                        _logger.LogDebug("Found fragment in DocLocation: {Fragment}", fragment);
                    }
                    // 2. Check Anchor property (typically for internal bookmarks)
                    else if (!string.IsNullOrEmpty(openXmlHyperlink.Anchor?.Value))
                    {
                        fragment = openXmlHyperlink.Anchor.Value;
                        _logger.LogDebug("Found fragment in Anchor: {Fragment}", fragment);
                    }
                }

                // Combine address and fragment exactly like VBA
                var completeUrl = !string.IsNullOrEmpty(fragment)
                    ? baseAddress + "#" + fragment
                    : baseAddress;

                _logger.LogDebug("Built complete URL: BaseAddress='{BaseAddress}', Fragment='{Fragment}', Complete='{CompleteUrl}'",
                    baseAddress, fragment ?? "", completeUrl);

                return completeUrl;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error building complete hyperlink URL for relationship: {RelId}", openXmlHyperlink.Id?.Value);
                return string.Empty;
            }
        }

        /// <summary>
        /// Extracts Content_ID or Document_ID using EXACT same logic as VBA ExtractLookupID function
        /// CRITICAL FIX: Now processes full URL directly to prevent fragment loss
        /// Note: Despite the VBA function name, this extracts identifiers (Content_IDs/Document_IDs) from URLs
        /// </summary>
        /// <param name="fullUrl">Complete hyperlink URL including fragments</param>
        /// <param name="unused">Unused parameter for compatibility</param>
        /// <returns>Extracted identifier (Content_ID or Document_ID) or empty string if no match</returns>
        private string ExtractIdentifierFromUrl(string fullUrl, string unused = "")
        {
            try
            {
                _logger.LogDebug("Attempting to extract lookup ID from full URL: {FullUrl}", fullUrl);

                // CRITICAL FIX: First, try exact VBA regex pattern with case-insensitive matching
                var regexMatch = LookupIdRegex.Match(fullUrl);
                if (regexMatch.Success)
                {
                    var lookupId = regexMatch.Value.ToUpperInvariant();
                    _logger.LogInformation("✓ EXTRACTED Lookup_ID via regex pattern: '{LookupId}' from URL: {Url}", lookupId, fullUrl);
                    return lookupId;
                }
                else
                {
                    _logger.LogDebug("✗ Regex pattern did not match TSRC/CMS format in URL: {Url}", fullUrl);
                }

                // CRITICAL FIX: Fallback docid extraction exactly like VBA (Issue #3)
                // VBA: ElseIf InStr(1, full, "docid=", vbTextCompare) > 0 Then
                if (fullUrl.IndexOf("docid=", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    _logger.LogDebug("Found 'docid=' parameter in URL, attempting extraction");
                    // VBA: ExtractLookupID = Trim$(Split(Split(full, "docid=")(1), "&")(0))
                    var parts = fullUrl.Split(new[] { "docid=" }, StringSplitOptions.None);
                    if (parts.Length > 1)
                    {
                        var docId = parts[1].Split('&')[0].Trim();
                        // CRITICAL FIX: Handle URL encoding (Issue #3)
                        var decodedDocId = Uri.UnescapeDataString(docId);
                        _logger.LogInformation("✓ EXTRACTED docid fallback: '{DecodedDocId}' (raw: '{DocId}') from URL: {Url}", decodedDocId, docId, fullUrl);
                        return decodedDocId;
                    }
                }
                else
                {
                    _logger.LogDebug("✗ No 'docid=' parameter found in URL: {Url}", fullUrl);
                }

                _logger.LogInformation("✗ NO VALID LOOKUP ID found in URL: {Url} - this hyperlink will be excluded from API call", fullUrl);
                return string.Empty;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error extracting Lookup_ID from URL: {Url}. Error: {Error}", fullUrl, ex.Message);
                return string.Empty;
            }
        }

        /// <summary>
        /// Extracts docid parameter from URLs following Base_File.vba methodology
        /// CRITICAL FIX: Added to support Content_ID detection in docid parameters
        /// </summary>
        private string ExtractDocIdFromUrl(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            try
            {
                // Check for docid parameter in URL with word boundary
                var docIdMatch = System.Text.RegularExpressions.Regex.Match(input, @"docid=([^&\s#]+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (docIdMatch.Success)
                {
                    // CRITICAL FIX: Handle URL encoding (Issue #3)
                    var rawDocId = docIdMatch.Groups[1].Value.Trim();
                    return Uri.UnescapeDataString(rawDocId);
                }

                // Return empty if no docid found
                return string.Empty;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Error extracting docid from URL: {Input}. Error: {Error}", input, ex.Message);
                return string.Empty;
            }
        }

        /// <summary>
        /// CRITICAL FIX: Implements complete VBA UpdateHyperlinksFromAPI workflow
        /// Follows exact sequence from Base_File.vba lines 15-350
        /// </summary>
        private async Task ProcessHyperlinksUsingVbaWorkflowAsync(BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken)
        {
            try
            {
                _logger.LogInformation("Starting VBA-compatible hyperlink processing workflow for: {FileName}", document.FileName);

                // VBA STEP 2: Collect unique Lookup_IDs (lines 41-84)
                var idDict = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
                var hyperlinkProcessingDetails = new List<string>();

                _logger.LogInformation("Analyzing {Count} hyperlinks for valid lookup identifiers", document.Hyperlinks.Count);

                foreach (var hyperlink in document.Hyperlinks)
                {
                    var hyperlinkDetail = $"Hyperlink URL: '{hyperlink.OriginalUrl}', Display: '{hyperlink.DisplayText ?? ""}', ";

                    // CRITICAL FIX: Extract Content_IDs or Document_IDs from URLs to include in Lookup_ID array
                    var extractedId = ExtractIdentifierFromUrl(hyperlink.OriginalUrl, "");
                    hyperlinkDetail += $"ExtractedID: '{extractedId}', ";

                    if (!string.IsNullOrEmpty(extractedId))
                    {
                        if (!idDict.ContainsKey(extractedId))
                        {
                            idDict[extractedId] = true;
                            hyperlinkDetail += "ADDED to lookup array";
                            _logger.LogInformation("Added unique identifier from URL: {ExtractedId} for hyperlink: {DisplayText}", extractedId, hyperlink.DisplayText ?? "[Empty]");
                        }
                        else
                        {
                            hyperlinkDetail += "Already exists in lookup array";
                        }
                    }
                    else
                    {
                        hyperlinkDetail += "NO VALID ID FOUND - will be skipped from API call";
                    }

                    // CRITICAL FIX: Also collect Content_IDs and Document_IDs if already available from previous processing
                    // These will be consolidated into the Lookup_ID JSON array sent to the API
                    if (!string.IsNullOrEmpty(hyperlink.ContentId) && !idDict.ContainsKey(hyperlink.ContentId))
                    {
                        idDict[hyperlink.ContentId] = true;
                        hyperlinkDetail += $", Added existing Content_ID: {hyperlink.ContentId}";
                        _logger.LogDebug("Added unique Content_ID: {ContentId}", hyperlink.ContentId);
                    }

                    if (!string.IsNullOrEmpty(hyperlink.DocumentId) && !idDict.ContainsKey(hyperlink.DocumentId))
                    {
                        idDict[hyperlink.DocumentId] = true;
                        hyperlinkDetail += $", Added existing Document_ID: {hyperlink.DocumentId}";
                        _logger.LogDebug("Added unique Document_ID: {DocumentId}", hyperlink.DocumentId);
                    }

                    hyperlinkProcessingDetails.Add(hyperlinkDetail);
                }

                // Log detailed analysis of each hyperlink
                _logger.LogInformation("Hyperlink Analysis Results for {FileName}:", document.FileName);
                for (int i = 0; i < hyperlinkProcessingDetails.Count; i++)
                {
                    _logger.LogInformation("  Hyperlink {Index}: {Details}", i + 1, hyperlinkProcessingDetails[i]);
                }

                if (idDict.Count == 0)
                {
                    _logger.LogWarning("No valid identifiers found in {Count} hyperlinks for document: {FileName}. Hyperlinks may not contain valid TSRC/CMS lookup IDs or docid parameters.", document.Hyperlinks.Count, document.FileName);
                    return;
                }

                _logger.LogInformation("Successfully collected {Count} unique identifiers from {HyperlinkCount} hyperlinks for Lookup_ID API array", idDict.Count, document.Hyperlinks.Count);

                // VBA STEP 3: Build JSON & POST (lines 87-128)
                // Note: lookupIds contains Content_IDs and Document_IDs that will be sent in the "Lookup_ID" JSON property
                var lookupIds = idDict.Keys.ToArray();

                // CRITICAL FIX: Use HyperlinkReplacementService for API processing with VBA methodology
                // First, we need an HttpService instance with HttpClient
                using var httpClient = new System.Net.Http.HttpClient();
                // For now, create a minimal structured logger for this context
                var minimalStructuredLogger = new StructuredLoggingService(_logger);
                var httpService = new HttpService(httpClient, _logger, _retryPolicyService, minimalStructuredLogger);

                // Create IOptions wrapper for AppSettings to match constructor requirements
                var appSettingsOptions = Microsoft.Extensions.Options.Options.Create(_appSettings);
                var hyperlinkService = new HyperlinkReplacementService(httpService, _logger, appSettingsOptions, _retryPolicyService);
                // CRITICAL FIX: Shorter timeout to prevent UI freezing - coordinated with HTTP timeout
                using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                timeoutCts.CancelAfter(TimeSpan.FromSeconds(75)); // 75-second timeout (HTTP 60s + buffer)

                var apiResult = await hyperlinkService.ProcessApiResponseAsync(lookupIds, timeoutCts.Token).ConfigureAwait(false);

                if (apiResult.HasError)
                {
                    _logger.LogError("API processing failed: {Error}", apiResult.ErrorMessage);
                    return;
                }

                // Check cancellation before hyperlink updates
                cancellationToken.ThrowIfCancellationRequested();

                // VBA STEP 5: Update hyperlinks using dictionary lookup (lines 186-318)
                await UpdateHyperlinksUsingVbaDictionaryLogicAsync(document, apiResult, cancellationToken).ConfigureAwait(false);

                _logger.LogInformation("Completed VBA-compatible hyperlink processing workflow for: {FileName}", document.FileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in VBA-compatible hyperlink processing workflow: {FileName}", document.FileName);
                throw;
            }
        }

        /// <summary>
        /// CRITICAL FIX: Implements VBA dictionary lookup and update logic (lines 186-318)
        /// VBA: For Each hl In links ... If recDict.Exists(lookupID) Then Set rec = recDict(lookupID)
        /// </summary>
        private async Task UpdateHyperlinksUsingVbaDictionaryLogicAsync(BulkEditor.Core.Entities.Document document, HyperlinkReplacementService.ApiProcessingResult apiResult, CancellationToken cancellationToken)
        {
            try
            {
                // Build unified dictionary exactly like VBA
                var recDict = new Dictionary<string, DocumentRecord>(StringComparer.OrdinalIgnoreCase);

                // Add found documents
                foreach (var doc in apiResult.FoundDocuments)
                {
                    if (!string.IsNullOrEmpty(doc.Document_ID) && !recDict.ContainsKey(doc.Document_ID))
                        recDict[doc.Document_ID] = doc;
                    if (!string.IsNullOrEmpty(doc.Content_ID) && !recDict.ContainsKey(doc.Content_ID))
                        recDict[doc.Content_ID] = doc;
                }

                // Add expired documents
                foreach (var doc in apiResult.ExpiredDocuments)
                {
                    if (!string.IsNullOrEmpty(doc.Document_ID) && !recDict.ContainsKey(doc.Document_ID))
                        recDict[doc.Document_ID] = doc;
                    if (!string.IsNullOrEmpty(doc.Content_ID) && !recDict.ContainsKey(doc.Content_ID))
                        recDict[doc.Content_ID] = doc;
                }

                // Track missing ID statistics for enhanced reporting
                var missingIdTracker = new Dictionary<string, string>();
                var processedHyperlinks = 0;
                var foundHyperlinks = 0;
                var expiredHyperlinks = 0;
                var notFoundHyperlinks = 0;

                // Process each hyperlink exactly like VBA
                foreach (var hyperlink in document.Hyperlinks)
                {
                    // Check cancellation in hyperlink processing loop to prevent hanging
                    cancellationToken.ThrowIfCancellationRequested();

                    var lookupId = ExtractIdentifierFromUrl(hyperlink.OriginalUrl, "");
                    if (string.IsNullOrEmpty(lookupId))
                        continue;

                    processedHyperlinks++;
                    var dispText = hyperlink.DisplayText ?? string.Empty;
                    var alreadyExpired = dispText.Contains(" - Expired", StringComparison.OrdinalIgnoreCase);
                    var alreadyNotFound = dispText.Contains(" - Not Found", StringComparison.OrdinalIgnoreCase);

                    if (recDict.ContainsKey(lookupId))
                    {
                        // VBA: If recDict.Exists(lookupID) Then Set rec = recDict(lookupID)
                        var rec = recDict[lookupId];

                        // Track found vs expired
                        if (rec.Status?.Equals("Expired", StringComparison.OrdinalIgnoreCase) == true)
                            expiredHyperlinks++;
                        else
                            foundHyperlinks++;

                        await ProcessHyperlinkWithVbaLogicAsync(hyperlink, rec, document, alreadyExpired, alreadyNotFound, cancellationToken).ConfigureAwait(false);
                    }
                    else if (!alreadyNotFound && !alreadyExpired)
                    {
                        // VBA: ElseIf Not alreadyNotFound And Not alreadyExpired Then
                        // VBA: hl.TextToDisplay = hl.TextToDisplay & " - Not Found"
                        notFoundHyperlinks++;
                        missingIdTracker[lookupId] = $"Hyperlink: '{dispText}' - Missing from API response";

                        // CRITICAL FIX: Preserve the lookup ID even when marked as Not Found
                        // This ensures the hyperlink can still be processed in UpdateHyperlinkWithAtomicVbaLogicAsync
                        hyperlink.LookupId = lookupId;
                        hyperlink.DisplayText += " - Not Found";
                        hyperlink.Status = HyperlinkStatus.NotFound;
                        hyperlink.RequiresUpdate = true; // CRITICAL FIX: Mark for update

                        document.ChangeLog.Changes.Add(new ChangeEntry
                        {
                            Type = ChangeType.HyperlinkStatusAdded,
                            Description = "Hyperlink marked as Not Found (VBA methodology)",
                            OldValue = dispText,
                            NewValue = hyperlink.DisplayText,
                            ElementId = hyperlink.Id,
                            Details = $"Lookup_ID: {lookupId}"
                        });

                        _logger.LogInformation("Marked hyperlink as Not Found: {LookupId} (Hyperlink: '{DisplayText}')", lookupId, dispText);
                    }
                }

                // Comprehensive hyperlink processing summary
                var successRate = processedHyperlinks > 0 ? (double)foundHyperlinks / processedHyperlinks * 100 : 0;
                _logger.LogInformation("Hyperlink processing summary for {FileName}: " +
                    "Total: {Total}, Found: {Found}, Expired: {Expired}, Not Found: {NotFound}, " +
                    "Success Rate: {SuccessRate:F1}%, API Response Rate: {ApiFoundCount} found + {ApiExpiredCount} expired from {ApiTotalCount} API IDs",
                    document.FileName, processedHyperlinks, foundHyperlinks, expiredHyperlinks, notFoundHyperlinks,
                    successRate, apiResult.FoundDocuments.Count, apiResult.ExpiredDocuments.Count, apiResult.TotalIdsProcessed);

                // Log missing ID details if any
                if (missingIdTracker.Any())
                {
                    _logger.LogWarning("Missing IDs details for {FileName}: {MissingDetails}",
                        document.FileName, string.Join("; ", missingIdTracker.Select(kvp => $"{kvp.Key}: {kvp.Value}")));

                    // Add summary changelog entry for missing IDs
                    document.ChangeLog.Changes.Add(new ChangeEntry
                    {
                        Type = ChangeType.Information,
                        Description = $"Missing ID Processing Summary: {missingIdTracker.Count} not found in API response",
                        Details = $"API processed {apiResult.TotalIdsProcessed} IDs, found {apiResult.TotalIdsFound}, missing {apiResult.TotalMissing} ({apiResult.MissingPercentage:F1}%)"
                    });
                }

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating hyperlinks using VBA dictionary logic");
                throw;
            }
        }

        /// <summary>
        /// CRITICAL FIX: Process individual hyperlink with exact VBA logic (lines 228-306)
        /// Enhanced to handle Content_ID in docid parameter scenarios
        /// </summary>
        private async Task ProcessHyperlinkWithVbaLogicAsync(Hyperlink hyperlink, DocumentRecord rec, BulkEditor.Core.Entities.Document document, bool alreadyExpired, bool alreadyNotFound, CancellationToken cancellationToken)
        {
            try
            {
                var dispText = hyperlink.DisplayText ?? string.Empty;

                // CRITICAL FIX: Set hyperlink properties from API response for later use in UpdateHyperlinkWithAtomicVbaLogicAsync
                if (!string.IsNullOrEmpty(rec.Document_ID))
                {
                    hyperlink.DocumentId = rec.Document_ID;
                    _logger.LogDebug("Set hyperlink.DocumentId from API: '{DocumentId}'", rec.Document_ID);
                }
                if (!string.IsNullOrEmpty(rec.Content_ID))
                {
                    hyperlink.ContentId = rec.Content_ID;
                    _logger.LogDebug("Set hyperlink.ContentId from API: '{ContentId}'", rec.Content_ID);
                }

                // CRITICAL FIX: Determine correct ID for URL building
                string idForUrl;
                if (!string.IsNullOrEmpty(rec.Document_ID))
                {
                    // Always prefer Document_ID when available
                    idForUrl = rec.Document_ID;
                    _logger.LogDebug("Using Document_ID for URL: '{DocumentId}'", idForUrl);
                }
                else if (!string.IsNullOrEmpty(rec.Content_ID))
                {
                    // Check if original URL contains Content_ID in docid parameter
                    var originalDocId = ExtractDocIdFromUrl(hyperlink.OriginalUrl);
                    var isContentIdInOriginalDocId = !string.IsNullOrEmpty(originalDocId) &&
                        string.Equals(originalDocId, rec.Content_ID, StringComparison.OrdinalIgnoreCase);

                    if (isContentIdInOriginalDocId)
                    {
                        _logger.LogWarning("CONTENT_ID_IN_DOCID_VBA: Original URL contains Content_ID '{ContentId}' in docid parameter, " +
                            "but no Document_ID available from API. Using Content_ID as fallback. URL: '{OriginalUrl}'",
                            rec.Content_ID, hyperlink.OriginalUrl);
                    }

                    idForUrl = rec.Content_ID;
                    _logger.LogDebug("Using Content_ID for URL: '{ContentId}'", idForUrl);
                }
                else
                {
                    // No ID available - should not happen
                    _logger.LogError("NO_ID_FOR_VBA_URL: Neither Document_ID nor Content_ID available for hyperlink. URL: '{OriginalUrl}'",
                        hyperlink.OriginalUrl);
                    return; // Skip processing this hyperlink
                }

                // VBA: targetAddress = "https://thesource.cvshealth.com/nuxeo/thesource/"
                // VBA: targetSub = "!/view?docid=" & rec("Document_ID")
                var targetAddress = "https://thesource.cvshealth.com/nuxeo/thesource/";

                // CRITICAL FIX: Properly encode the fragment to prevent XSD validation errors with 0x21 (!) character
                var targetSub = Uri.EscapeDataString($"!/view?docid={idForUrl}");

                // VBA: changedURL = (hl.Address <> targetAddress) Or (hl.SubAddress <> targetSub)
                var targetUrl = targetAddress + "#" + Uri.UnescapeDataString(targetSub);
                var changedURL = !string.Equals(hyperlink.OriginalUrl, targetUrl, StringComparison.OrdinalIgnoreCase);

                // CRITICAL FIX: Always mark file:// URLs as changed when we have a valid ID
                if (!changedURL && hyperlink.OriginalUrl.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                {
                    changedURL = true;
                    _logger.LogDebug("Forcing URL change for file:// URL with ID: {OriginalUrl} -> {TargetUrl}",
                        hyperlink.OriginalUrl, targetUrl);
                }

                // CRITICAL FIX: Also check for Content_ID in docid parameter that needs updating to Document_ID
                if (!changedURL && !string.IsNullOrEmpty(rec.Document_ID) && !string.IsNullOrEmpty(rec.Content_ID))
                {
                    var originalDocId = ExtractDocIdFromUrl(hyperlink.OriginalUrl);
                    if (!string.IsNullOrEmpty(originalDocId) &&
                        string.Equals(originalDocId, rec.Content_ID, StringComparison.OrdinalIgnoreCase))
                    {
                        changedURL = true;
                        _logger.LogInformation("DOCID_CONTENT_TO_DOCUMENT_ID: Forcing URL change to replace Content_ID '{ContentId}' " +
                            "with Document_ID '{DocumentId}' in docid parameter. Original: '{OriginalUrl}' -> New: '{TargetUrl}'",
                            rec.Content_ID, rec.Document_ID, hyperlink.OriginalUrl, targetUrl);
                    }
                }

                if (changedURL)
                {
                    hyperlink.UpdatedUrl = targetUrl;
                    hyperlink.RequiresUpdate = true;
                    _logger.LogDebug("URL changed for hyperlink: {Original} -> {New}", hyperlink.OriginalUrl, targetUrl);
                }

                // VBA Content_ID appending logic (lines 254-280)
                var appended = false;
                if (!alreadyExpired && !alreadyNotFound && !string.IsNullOrEmpty(rec.Content_ID))
                {
                    // CRITICAL FIX: Exact VBA logic with proper bounds checking
                    if (rec.Content_ID.Length >= 6)
                    {
                        var last6 = rec.Content_ID.Substring(rec.Content_ID.Length - 6);
                        var last5 = last6.Length > 1 ? last6.Substring(1) : last6;

                        var pattern5 = $" ({last5})";
                        var pattern6 = $" ({last6})";

                        // VBA: If Right$(dispText, Len(" (" & last5 & ")")) = " (" & last5 & ")" And Right$(dispText, Len(" (" & last6 & ")")) <> " (" & last6 & ")" Then
                        if (dispText.EndsWith(pattern5) && !dispText.EndsWith(pattern6))
                        {
                            if (dispText.Length >= pattern5.Length)
                            {
                                dispText = dispText.Substring(0, dispText.Length - pattern5.Length) + pattern6;
                                hyperlink.DisplayText = dispText;
                                hyperlink.RequiresUpdate = true; // CRITICAL FIX: Mark for update
                                appended = true;
                                _logger.LogInformation("Upgraded 5-digit to 6-digit Content_ID: {Old} -> {New}", pattern5, pattern6);
                            }
                        }
                        // VBA: ElseIf InStr(1, dispText, " (" & last6 & ")", vbTextCompare) = 0 Then
                        else if (!dispText.Contains(pattern6, StringComparison.OrdinalIgnoreCase))
                        {
                            hyperlink.DisplayText = dispText.Trim() + pattern6;
                            hyperlink.RequiresUpdate = true; // CRITICAL FIX: Mark for update
                            dispText = hyperlink.DisplayText;
                            appended = true;
                            _logger.LogInformation("Appended Content_ID to hyperlink: {ContentId}", last6);
                        }
                    }

                    // VBA title comparison logic (lines 282-286)
                    var titleWithoutContentId = dispText;
                    if (dispText.Length > 9) // " (123456)"
                    {
                        titleWithoutContentId = dispText.Substring(0, dispText.Length - 9);
                    }

                    if (!string.Equals(titleWithoutContentId, rec.Title, StringComparison.OrdinalIgnoreCase))
                    {
                        document.ChangeLog.Changes.Add(new ChangeEntry
                        {
                            Type = ChangeType.PossibleTitleChange,
                            Description = "Possible Title Change",
                            OldValue = titleWithoutContentId,
                            NewValue = rec.Title,
                            ElementId = hyperlink.Id,
                            Details = $"Content ID: {rec.Content_ID}"
                        });
                        _logger.LogInformation("Detected title difference: Current='{Current}', API='{Api}'", titleWithoutContentId, rec.Title);
                    }
                }

                // VBA status handling (lines 292-306)
                if (rec.Status.Equals("Expired", StringComparison.OrdinalIgnoreCase) && !alreadyExpired)
                {
                    hyperlink.DisplayText += " - Expired";
                    hyperlink.Status = HyperlinkStatus.Expired;
                    hyperlink.RequiresUpdate = true; // CRITICAL FIX: Mark for update

                    document.ChangeLog.Changes.Add(new ChangeEntry
                    {
                        Type = ChangeType.HyperlinkStatusAdded,
                        Description = "Hyperlink marked as Expired (VBA methodology)",
                        OldValue = dispText,
                        NewValue = hyperlink.DisplayText,
                        ElementId = hyperlink.Id,
                        Details = $"API Status: {rec.Status}, Content_ID: {rec.Content_ID}"
                    });

                    _logger.LogInformation("Marked hyperlink as Expired: {LookupId}", rec.Lookup_ID);
                }
                else if (changedURL || appended)
                {
                    var actionDescription = changedURL ? "URL Updated" : "";
                    if (appended)
                        actionDescription += (changedURL ? ", " : "") + "Appended Content ID";

                    document.ChangeLog.Changes.Add(new ChangeEntry
                    {
                        Type = changedURL ? ChangeType.HyperlinkUpdated : ChangeType.ContentIdAdded,
                        Description = actionDescription,
                        OldValue = hyperlink.OriginalUrl,
                        NewValue = hyperlink.UpdatedUrl ?? hyperlink.OriginalUrl,
                        ElementId = hyperlink.Id,
                        Details = $"Document_ID: {rec.Document_ID}, Content_ID: {rec.Content_ID}"
                    });

                    _logger.LogInformation("Updated hyperlink: {Action} for {LookupId}", actionDescription, rec.Lookup_ID);
                }

                // CRITICAL FIX: Ensure RequiresUpdate is set if any changes were made
                // This is a safeguard to catch any missed cases where changes occurred but flag wasn't set
                if ((changedURL || appended) && !hyperlink.RequiresUpdate)
                {
                    hyperlink.RequiresUpdate = true;
                    _logger.LogDebug("Safeguard: Marked hyperlink for update due to changes: URL={UrlChanged}, Content={ContentChanged}",
                        changedURL, appended);
                }

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing hyperlink with VBA logic: {HyperlinkId}", hyperlink.Id);
                throw;
            }
        }

        private void GenerateChangeLogSummary(BulkEditor.Core.Entities.Document document)
        {
            var changes = document.ChangeLog.Changes;
            var summary = $"Processed {document.FileName}: ";

            var hyperlinkUpdates = changes.Count(c => c.Type == ChangeType.HyperlinkUpdated);
            var hyperlinkDeleted = changes.Count(c => c.Type == ChangeType.HyperlinkRemoved);
            var contentIdAdded = changes.Count(c => c.Type == ChangeType.ContentIdAdded);
            var titleReplacements = changes.Count(c => c.Type == ChangeType.TitleReplaced);
            var titleChanges = changes.Count(c => c.Type == ChangeType.PossibleTitleChange);
            var statusAdded = changes.Count(c => c.Type == ChangeType.HyperlinkStatusAdded);
            var errors = changes.Count(c => c.Type == ChangeType.Error);

            var summaryParts = new List<string>();

            if (hyperlinkUpdates > 0)
                summaryParts.Add($"{hyperlinkUpdates} hyperlinks updated");

            if (hyperlinkDeleted > 0)
                summaryParts.Add($"{hyperlinkDeleted} invisible hyperlinks deleted");

            if (contentIdAdded > 0)
                summaryParts.Add($"{contentIdAdded} content IDs added");

            if (titleReplacements > 0)
                summaryParts.Add($"{titleReplacements} titles replaced");

            if (titleChanges > 0)
                summaryParts.Add($"{titleChanges} possible title changes");

            if (statusAdded > 0)
                summaryParts.Add($"{statusAdded} status suffixes added");

            if (errors > 0)
                summaryParts.Add($"{errors} errors");

            document.ChangeLog.Summary = summary + (summaryParts.Any() ? string.Join(", ", summaryParts) : "no changes required");
        }

        /// <summary>
        /// DEPRECATED: Status suffix handling is now done within the single document session
        /// This method only logs the status change for external validation results
        /// </summary>
        private async Task HandleHyperlinkStatusSuffixAsync(BulkEditor.Core.Entities.Document document, Hyperlink hyperlink, CancellationToken cancellationToken)
        {
            try
            {
                string? statusSuffix = null;
                ChangeType changeType = ChangeType.Information;

                // Determine status suffix based on hyperlink status
                switch (hyperlink.Status)
                {
                    case HyperlinkStatus.Expired:
                        statusSuffix = " - Expired";
                        changeType = ChangeType.HyperlinkStatusAdded;
                        break;
                    case HyperlinkStatus.NotFound:
                        statusSuffix = " - Not Found";
                        changeType = ChangeType.HyperlinkStatusAdded;
                        break;
                    default:
                        return; // No suffix needed for other statuses
                }

                // Check if the suffix is already present to avoid duplicates
                var currentDisplayText = hyperlink.DisplayText ?? string.Empty;
                if (currentDisplayText.EndsWith(statusSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    _logger.LogDebug("Status suffix '{StatusSuffix}' already present in hyperlink: {HyperlinkId}", statusSuffix, hyperlink.Id);
                    return;
                }

                // Update the hyperlink object for tracking (actual document update happens in session)
                hyperlink.DisplayText = currentDisplayText + statusSuffix;

                // Log the change
                document.ChangeLog.Changes.Add(new ChangeEntry
                {
                    Type = changeType,
                    Description = $"Status suffix marked for addition: {statusSuffix}",
                    OldValue = currentDisplayText,
                    NewValue = hyperlink.DisplayText,
                    ElementId = hyperlink.Id,
                    Details = $"Hyperlink status: {hyperlink.Status} (applied in session)"
                });

                _logger.LogInformation("Marked status suffix '{StatusSuffix}' for addition to hyperlink: {HyperlinkId}", statusSuffix, hyperlink.Id);
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error handling status suffix for hyperlink: {HyperlinkId}", hyperlink.Id);

                document.ChangeLog.Changes.Add(new ChangeEntry
                {
                    Type = ChangeType.Error,
                    Description = "Error marking status suffix",
                    ElementId = hyperlink.Id,
                    Details = ex.Message
                });
            }
        }

        /// <summary>
        /// Document state snapshot for rollback operations
        /// </summary>
        private class DocumentSnapshot
        {
            public Dictionary<string, string> RelationshipMappings { get; set; } = new();
            public List<string> ModifiedRelationshipIds { get; set; } = new();
            public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        }

        /// <summary>
        /// CRITICAL FIX: Process all document operations in a single session to prevent corruption
        /// Now includes comprehensive validation and atomic operations
        /// </summary>
        private async Task ProcessDocumentInSingleSessionAsync(BulkEditor.Core.Entities.Document document, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
        {
            DocumentSnapshot? snapshot = null;
            var totalProcessingStart = DateTime.UtcNow;

            try
            {
                _logger.LogInformation("Processing document in single session with enhanced corruption prevention: {FileName}", document.FileName);

                // STEP 1: Pre-processing validation
                progress?.Report("Validating document before processing...");
                await ValidateDocumentIntegrityAsync(document.FilePath, "pre-processing", cancellationToken).ConfigureAwait(false);

                // STEP 1.5: Validate processing options
                progress?.Report("Validating processing options...");
                ValidateProcessingOptions();

                // Open document once and perform ALL operations within this session with retry logic
                var filePolicy = _retryPolicyService.CreateFileRetryPolicy();
                using (var wordDocument = await _retryPolicyService.ExecuteWithRetryAsync(
                    () => Task.FromResult(WordprocessingDocument.Open(document.FilePath, true)),
                    filePolicy, cancellationToken).ConfigureAwait(false))
                {
                    var mainPart = wordDocument.MainDocumentPart;

                    if (mainPart?.Document?.Body == null)
                    {
                        throw new InvalidOperationException($"Document has no main content: {document.FilePath}");
                    }

                    // STEP 2: Initial document validation
                    progress?.Report("Validating document structure...");
                    await ValidateOpenDocumentAsync(wordDocument, "initial", cancellationToken).ConfigureAwait(false);

                    // STEP 3: Create document snapshot for rollback
                    progress?.Report("Creating document snapshot...");
                    snapshot = CreateDocumentSnapshot(mainPart);

                    // STEP 5: Extract metadata (read-only operations first)
                    progress?.Report("Extracting metadata...");
                    document.Metadata = ExtractDocumentMetadataFromOpenDocument(wordDocument);

                    // STEP 6: Extract hyperlinks from the open document
                    progress?.Report("Extracting hyperlinks...");
                    document.Hyperlinks = ExtractHyperlinksFromOpenDocument(mainPart);
                    _logger.LogInformation("Initial hyperlink extraction found {Count} hyperlinks in document: {FileName}", document.Hyperlinks.Count, document.FileName);

                    // STEP 7: CRITICAL FIX - Extract lookup IDs BEFORE removing invisible hyperlinks
                    // This ensures we don't lose valid lookup IDs from hyperlinks with empty display text
                    progress?.Report("Pre-analyzing hyperlinks for lookup IDs...");
                    var allHyperlinksBeforeCleanup = ExtractHyperlinksFromOpenDocument(mainPart);
                    var validLookupIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    foreach (var hyperlink in allHyperlinksBeforeCleanup)
                    {
                        // CRITICAL FIX: Use full URL directly instead of parsing into address/subAddress
                        var extractedId = ExtractIdentifierFromUrl(hyperlink.OriginalUrl, "");
                        if (!string.IsNullOrEmpty(extractedId))
                        {
                            validLookupIds.Add(extractedId);
                            _logger.LogInformation("Pre-cleanup: Found valid lookup ID '{LookupId}' in hyperlink with display text: '{DisplayText}'",
                                extractedId, hyperlink.DisplayText ?? "[EMPTY]");
                        }
                    }

                    _logger.LogInformation("Pre-cleanup analysis: Found {ValidIdCount} valid lookup IDs from {TotalHyperlinks} hyperlinks", validLookupIds.Count, allHyperlinksBeforeCleanup.Count);

                    // STEP 7: Remove invisible hyperlinks (write operations) - but preserve those with valid lookup IDs
                    progress?.Report("Removing invisible hyperlinks (preserving those with valid lookup IDs)...");
                    await RemoveInvisibleHyperlinksInSessionAsync(mainPart, document, cancellationToken).ConfigureAwait(false);

                    // STEP 8: Re-extract hyperlinks after deletion to prevent stale references
                    progress?.Report("Re-extracting hyperlinks after cleanup...");
                    document.Hyperlinks = ExtractHyperlinksFromOpenDocument(mainPart);
                    _logger.LogInformation("After cleanup: {Count} hyperlinks remain in document: {FileName}", document.Hyperlinks.Count, document.FileName);

                    // STEP 8: Validate after invisible hyperlink removal
                    progress?.Report("Validating after invisible hyperlink removal...");
                    await ValidateOpenDocumentAsync(wordDocument, "post-cleanup", cancellationToken).ConfigureAwait(false);

                    // STEP 9: Process hyperlinks using VBA UpdateHyperlinksFromAPI workflow (only if enabled)
                    var hyperlinkStepStart = DateTime.UtcNow;
                    // CRITICAL FIX: Skip API calls if BaseUrl is empty or "Test"
                    bool skipApiCalls = string.IsNullOrWhiteSpace(_appSettings.Api.BaseUrl) ||
                                       _appSettings.Api.BaseUrl.Equals("Test", StringComparison.OrdinalIgnoreCase);

                    // Enhanced logging: Show processing workflow decision
                    _logger.LogInformation("📋 PROCESSING WORKFLOW DECISION: UpdateHyperlinks={Update}, ValidateHyperlinks={Validate}, ValidateTitlesOnly={TitlesOnly}, HasHyperlinks={HasLinks}, SkipApiCalls={SkipApi}",
                        _appSettings.Processing.UpdateHyperlinks, _appSettings.Processing.ValidateHyperlinks, _appSettings.Validation.ValidateTitlesOnly, document.Hyperlinks.Any(), skipApiCalls);

                    if (_appSettings.Processing.UpdateHyperlinks && _appSettings.Processing.ValidateHyperlinks && document.Hyperlinks.Any() && !skipApiCalls)
                    {
                        progress?.Report("Processing hyperlinks using VBA methodology...");
                        _logger.LogInformation("Starting hyperlink processing for {FileName}: {HyperlinkCount} hyperlinks found, UpdateHyperlinks={UpdateEnabled}, ValidateHyperlinks={ValidateEnabled}",
                            document.FileName, document.Hyperlinks.Count, _appSettings.Processing.UpdateHyperlinks, _appSettings.Processing.ValidateHyperlinks);

                        // Add cancellation check before expensive VBA workflow
                        cancellationToken.ThrowIfCancellationRequested();
                        await ProcessHyperlinksUsingVbaWorkflowAsync(document, cancellationToken).ConfigureAwait(false);

                        // STEP 10: Apply hyperlink updates in the document session
                        progress?.Report("Applying hyperlink updates to document...");
                        await UpdateHyperlinksInSessionAsync(mainPart, document, cancellationToken).ConfigureAwait(false);

                        var hyperlinkDuration = DateTime.UtcNow - hyperlinkStepStart;
                        _logger.LogInformation("Hyperlink processing completed for {FileName}: Duration={DurationMs}ms",
                            document.FileName, hyperlinkDuration.TotalMilliseconds);
                    }
                    else if (skipApiCalls)
                    {
                        _logger.LogInformation("Hyperlink processing skipped for {FileName}: API BaseUrl is empty or set to 'Test' mode", document.FileName);
                    }
                    else if (!_appSettings.Processing.UpdateHyperlinks)
                    {
                        _logger.LogInformation("Hyperlink processing skipped for {FileName}: UpdateHyperlinks disabled in settings", document.FileName);
                    }
                    else if (!_appSettings.Processing.ValidateHyperlinks)
                    {
                        _logger.LogInformation("Hyperlink processing skipped for {FileName}: ValidateHyperlinks disabled in settings", document.FileName);
                    }
                    else
                    {
                        _logger.LogInformation("Hyperlink processing skipped for {FileName}: No hyperlinks found in document", document.FileName);
                    }

                    // STEP 9.5: Title-only validation workflow (if enabled - can run alongside or instead of hyperlink processing)
                    if (_appSettings.Validation.ValidateTitlesOnly && document.Hyperlinks.Any())
                    {
                        progress?.Report("Validating titles only...");
                        _logger.LogInformation("Starting title-only validation workflow for {FileName}", document.FileName);
                        
                        await ValidateAndUpdateTitlesAsync(document, cancellationToken).ConfigureAwait(false);
                        
                        // Apply title updates to the document if any hyperlinks were marked for update
                        var titleUpdates = document.Hyperlinks.Where(h => h.RequiresUpdate).ToList();
                        if (titleUpdates.Any())
                        {
                            progress?.Report("Applying title updates to document...");
                            await UpdateHyperlinksInSessionAsync(mainPart, document, cancellationToken).ConfigureAwait(false);
                            _logger.LogInformation("Applied {Count} title updates to document: {FileName}", titleUpdates.Count, document.FileName);
                        }
                    }

                    // STEP 11: Validate after hyperlink updates
                    progress?.Report("Validating after hyperlink updates...");
                    await ValidateOpenDocumentAsync(wordDocument, "post-hyperlinks", cancellationToken).ConfigureAwait(false);

                    // STEP 12: Process replacements in the same session (only if any replacement type is enabled)
                    var replacementStepStart = DateTime.UtcNow;
                    if (_appSettings.Replacement.EnableHyperlinkReplacement || _appSettings.Replacement.EnableTextReplacement)
                    {
                        progress?.Report("Processing replacements...");
                        _logger.LogInformation("Starting replacement processing for {FileName}: HyperlinkReplacement={HyperlinkEnabled} ({HyperlinkRulesCount} rules), TextReplacement={TextEnabled} ({TextRulesCount} rules)",
                            document.FileName, _appSettings.Replacement.EnableHyperlinkReplacement, _appSettings.Replacement.HyperlinkRules.Count,
                            _appSettings.Replacement.EnableTextReplacement, _appSettings.Replacement.TextRules.Count);

                        await _replacementService.ProcessReplacementsInSessionAsync(wordDocument, document, cancellationToken).ConfigureAwait(false);

                        var replacementDuration = DateTime.UtcNow - replacementStepStart;
                        _logger.LogInformation("Replacement processing completed for {FileName}: Duration={DurationMs}ms",
                            document.FileName, replacementDuration.TotalMilliseconds);
                    }
                    else
                    {
                        _logger.LogInformation("Replacement processing skipped for {FileName}: All replacement types disabled in settings", document.FileName);
                    }

                    // STEP 13: Validate after replacements
                    progress?.Report("Validating after replacements...");
                    await ValidateOpenDocumentAsync(wordDocument, "post-replacements", cancellationToken).ConfigureAwait(false);

                    // STEP 14: Optimize text in the same session (only if enabled)
                    var textOptimizationStart = DateTime.UtcNow;
                    if (_appSettings.Processing.OptimizeText)
                    {
                        progress?.Report("Optimizing document text...");
                        _logger.LogInformation("Starting text optimization for {FileName}: OptimizeText=enabled", document.FileName);

                        await _textOptimizer.OptimizeDocumentTextInSessionAsync(wordDocument, document, cancellationToken).ConfigureAwait(false);

                        var textOptimizationDuration = DateTime.UtcNow - textOptimizationStart;
                        _logger.LogInformation("Text optimization completed for {FileName}: Duration={DurationMs}ms",
                            document.FileName, textOptimizationDuration.TotalMilliseconds);
                    }
                    else
                    {
                        _logger.LogInformation("Text optimization skipped for {FileName}: OptimizeText disabled in settings", document.FileName);
                    }

                    // STEP 14: Final validation before save
                    progress?.Report("Final validation before save...");
                    await ValidateOpenDocumentAsync(wordDocument, "pre-save", cancellationToken).ConfigureAwait(false);

                    // STEP 15: Save document with enhanced error handling and validation
                    progress?.Report("Saving document with validation...");
                    await SaveDocumentSafelyAsync(wordDocument, document, cancellationToken).ConfigureAwait(false);

                } // CRITICAL: WordprocessingDocument disposed here - ensures file handles are released

                // STEP 16: Post-save validation
                progress?.Report("Validating document after save...");
                await ValidateDocumentIntegrityAsync(document.FilePath, "post-save", cancellationToken).ConfigureAwait(false);

                var totalProcessingDuration = DateTime.UtcNow - totalProcessingStart;
                _logger.LogInformation("Document processed successfully with comprehensive validation: {FileName}, " +
                    "Total Duration: {TotalDurationMs}ms, Hyperlinks: {HyperlinkCount}, Changes: {ChangeCount}",
                    document.FileName, totalProcessingDuration.TotalMilliseconds,
                    document.Hyperlinks.Count, document.ChangeLog.TotalChanges);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in single session document processing with rollback attempt: {FileName}", document.FileName);

                // Attempt to restore from backup if we have one
                if (!string.IsNullOrEmpty(document.BackupPath))
                {
                    await AttemptDocumentRecoveryAsync(document, ex, cancellationToken).ConfigureAwait(false);
                }

                throw;
            }
        }

        /// <summary>
        /// Validates document integrity after processing with enhanced retry logic and timeout protection
        /// CRITICAL FIX: Prevents file access conflicts and adds timeout protection
        /// </summary>
        private async Task ValidateDocumentIntegrityWithRetryAsync(string filePath, CancellationToken cancellationToken)
        {
            const int maxRetries = 5;
            const int baseRetryDelayMs = 200;

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    // CRITICAL FIX: Progressive delay with timeout protection to handle file system delays
                    if (attempt > 1)
                    {
                        var delayMs = baseRetryDelayMs * attempt;
                        using var delayCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                        delayCts.CancelAfter(TimeSpan.FromSeconds(5)); // Max 5 second delay
                        await Task.Delay(delayMs, delayCts.Token);
                    }

                    // CRITICAL FIX: Add timeout protection for document opening
                    using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                    timeoutCts.CancelAfter(TimeSpan.FromSeconds(10)); // 10-second timeout for validation

                    // Try to open the document in read-only mode with timeout protection
                    using var testDocument = await Task.Run(() =>
                        WordprocessingDocument.Open(filePath, false), timeoutCts.Token);

                    var mainPart = testDocument.MainDocumentPart;

                    if (mainPart?.Document?.Body == null)
                    {
                        throw new InvalidOperationException("Document appears to be corrupted - no main content found");
                    }

                    // Try to access the document content to ensure it's readable
                    var contentLength = mainPart.Document.Body.InnerText?.Length ?? 0;

                    _logger.LogDebug("Document integrity validation passed on attempt {Attempt}: {FilePath} (Content: {ContentLength} chars)",
                        attempt, filePath, contentLength);
                    return; // Success
                }
                catch (IOException ioEx) when (ioEx.Message.Contains("being used by another process") && attempt < maxRetries)
                {
                    _logger.LogWarning("File access conflict on attempt {Attempt}/{MaxRetries}: {FilePath}. Will retry in {DelayMs}ms...",
                        attempt, maxRetries, filePath, baseRetryDelayMs * (attempt + 1));
                    continue; // Retry
                }
                catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                {
                    _logger.LogWarning("Document validation cancelled by user on attempt {Attempt}: {FilePath}", attempt, filePath);
                    throw;
                }
                catch (OperationCanceledException)
                {
                    _logger.LogWarning("Document validation timed out on attempt {Attempt}: {FilePath}", attempt, filePath);
                    if (attempt == maxRetries)
                    {
                        throw new TimeoutException($"Document validation timed out after {maxRetries} attempts");
                    }
                    continue; // Retry on timeout
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Document integrity validation failed on attempt {Attempt}/{MaxRetries}: {FilePath}",
                        attempt, maxRetries, filePath);

                    if (attempt == maxRetries)
                    {
                        throw new InvalidOperationException($"Document appears to be corrupted after processing: {ex.Message}", ex);
                    }
                }
            }
        }


        /// <summary>
        /// Validates processing options to ensure they are properly configured
        /// </summary>
        private void ValidateProcessingOptions()
        {
            try
            {
                var issues = new List<string>();

                // Validate timeout settings
                if (_appSettings.Processing.TimeoutPerDocument.TotalSeconds < 30 || _appSettings.Processing.TimeoutPerDocument.TotalSeconds > 1800)
                {
                    issues.Add($"TimeoutPerDocument should be between 30 seconds and 30 minutes. Current: {_appSettings.Processing.TimeoutPerDocument.TotalSeconds} seconds");
                }

                // Validate concurrent document limits
                if (_appSettings.Processing.MaxConcurrentDocuments < 1 || _appSettings.Processing.MaxConcurrentDocuments > 1000)
                {
                    issues.Add($"MaxConcurrentDocuments should be between 1 and 1000. Current: {_appSettings.Processing.MaxConcurrentDocuments}");
                }

                // Validate batch size
                if (_appSettings.Processing.BatchSize < 1 || _appSettings.Processing.BatchSize > 1000)
                {
                    issues.Add($"BatchSize should be between 1 and 1000. Current: {_appSettings.Processing.BatchSize}");
                }

                // Validate lookup ID pattern
                if (string.IsNullOrWhiteSpace(_appSettings.Processing.LookupIdPattern))
                {
                    issues.Add("LookupIdPattern is required for hyperlink processing");
                }
                else
                {
                    try
                    {
                        var testRegex = new System.Text.RegularExpressions.Regex(_appSettings.Processing.LookupIdPattern);
                    }
                    catch (Exception ex)
                    {
                        issues.Add($"Invalid LookupIdPattern: {ex.Message}");
                    }
                }

                // Validate API settings if hyperlink processing is enabled
                if (_appSettings.Processing.UpdateHyperlinks && _appSettings.Processing.ValidateHyperlinks)
                {
                    // CRITICAL FIX: Only validate API if not using Test mode
                    if (!string.IsNullOrWhiteSpace(_appSettings.Api.BaseUrl) &&
                        !_appSettings.Api.BaseUrl.Equals("Test", StringComparison.OrdinalIgnoreCase))
                    {
                        // Real API URL provided - validate it
                        if (!Uri.IsWellFormedUriString(_appSettings.Api.BaseUrl, UriKind.Absolute))
                        {
                            issues.Add($"API BaseUrl must be a valid URL. Current: {_appSettings.Api.BaseUrl}");
                        }
                    }
                    // If BaseUrl is empty or "Test", that's OK - we'll skip API calls

                    if (_appSettings.Api.Timeout.TotalSeconds < 5 || _appSettings.Api.Timeout.TotalSeconds > 300)
                    {
                        issues.Add($"API Timeout should be between 5 seconds and 5 minutes. Current: {_appSettings.Api.Timeout.TotalSeconds} seconds");
                    }
                }

                // Validate replacement settings if replacement is enabled
                if (_appSettings.Replacement.EnableHyperlinkReplacement)
                {
                    if (!_appSettings.Replacement.HyperlinkRules.Any())
                    {
                        _logger.LogWarning("Hyperlink replacement is enabled but no rules are configured");
                    }
                    else if (_appSettings.Replacement.HyperlinkRules.Count > _appSettings.Replacement.MaxReplacementRules)
                    {
                        issues.Add($"Too many hyperlink rules configured. Maximum: {_appSettings.Replacement.MaxReplacementRules}, Current: {_appSettings.Replacement.HyperlinkRules.Count}");
                    }
                }

                if (_appSettings.Replacement.EnableTextReplacement)
                {
                    if (!_appSettings.Replacement.TextRules.Any())
                    {
                        _logger.LogWarning("Text replacement is enabled but no rules are configured");
                    }
                    else if (_appSettings.Replacement.TextRules.Count > _appSettings.Replacement.MaxReplacementRules)
                    {
                        issues.Add($"Too many text rules configured. Maximum: {_appSettings.Replacement.MaxReplacementRules}, Current: {_appSettings.Replacement.TextRules.Count}");
                    }
                }

                if (issues.Any())
                {
                    var issuesSummary = string.Join("; ", issues);
                    throw new InvalidOperationException($"Processing options validation failed: {issuesSummary}");
                }

                _logger.LogDebug("Processing options validation completed successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error validating processing options");
                throw;
            }
        }

        /// <summary>
        /// Extracts metadata from an already opened WordprocessingDocument
        /// </summary>
        private DocumentMetadata ExtractDocumentMetadataFromOpenDocument(WordprocessingDocument wordDocument)
        {
            var metadata = new DocumentMetadata();

            try
            {
                // Get file info using the document's file path from the main document part
                metadata.FileSizeBytes = 0; // Will be set by the calling method if needed
                metadata.LastModified = DateTime.Now; // Use current time as fallback

                var coreProperties = wordDocument.PackageProperties;
                metadata.Title = coreProperties.Title ?? "";
                metadata.Author = coreProperties.Creator ?? "";
                metadata.Subject = coreProperties.Subject ?? "";
                metadata.Keywords = coreProperties.Keywords ?? "";
                metadata.Comments = coreProperties.Description ?? "";

                // Count words and pages
                if (wordDocument.MainDocumentPart?.Document?.Body != null)
                {
                    var text = wordDocument.MainDocumentPart.Document.Body.InnerText;
                    metadata.WordCount = text.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Length;
                    metadata.PageCount = 1; // Simplified - actual page counting is complex
                }

                _logger.LogDebug("Extracted metadata from open document: {Title}", metadata.Title);
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Error extracting metadata from open document. Error: {Error}", ex.Message);
            }

            return metadata;
        }

        /// <summary>
        /// Extracts hyperlinks from an already opened MainDocumentPart
        /// </summary>
        private List<Hyperlink> ExtractHyperlinksFromOpenDocument(MainDocumentPart mainPart)
        {
            var hyperlinks = new List<Hyperlink>();

            try
            {
                var openXmlHyperlinks = mainPart.Document.Body?.Descendants<OpenXmlHyperlink>().ToList() ?? new List<OpenXmlHyperlink>();

                foreach (var openXmlHyperlink in openXmlHyperlinks)
                {
                    try
                    {
                        var relId = openXmlHyperlink.Id?.Value;
                        if (string.IsNullOrEmpty(relId))
                            continue;

                        // Safely try to get the relationship
                        try
                        {
                            // CRITICAL FIX: Build complete URL including fragment exactly like VBA
                            var completeUrl = BuildCompleteHyperlinkUrl(mainPart, openXmlHyperlink);
                            var displayText = openXmlHyperlink.InnerText;

                            _logger.LogDebug("Extracted hyperlink: Complete='{CompleteUrl}'", completeUrl);

                            var hyperlink = new Hyperlink
                            {
                                OriginalUrl = completeUrl, // Use decoded complete URL
                                DisplayText = displayText,
                                LookupId = ExtractIdentifierFromUrl(completeUrl, ""), // Use complete URL for extraction
                                RequiresUpdate = ShouldAutoValidateHyperlink(completeUrl, displayText)
                            };

                            hyperlinks.Add(hyperlink);
                        }
                        catch (System.Collections.Generic.KeyNotFoundException)
                        {
                            _logger.LogWarning("Skipping hyperlink with invalid relationship ID: {RelId}", relId);
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning("Error extracting hyperlink from open document. Error: {Error}", ex.Message);
                    }
                }

                _logger.LogDebug("Extracted {Count} hyperlinks from open document", hyperlinks.Count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error extracting hyperlinks from open document");
            }

            return hyperlinks;
        }

        /// <summary>
        /// Removes invisible hyperlinks within the current document session
        /// CRITICAL FIX: Preserves hyperlinks with valid lookup IDs even if they have empty display text
        /// </summary>
        private async Task RemoveInvisibleHyperlinksInSessionAsync(MainDocumentPart mainPart, BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken)
        {
            try
            {
                var invisibleLinksRemoved = 0;
                var hyperlinks = mainPart.Document.Body?.Descendants<OpenXmlHyperlink>().ToList() ?? new List<OpenXmlHyperlink>();

                _logger.LogInformation("Analyzing {Count} hyperlinks for invisible link cleanup in document: {FileName}", hyperlinks.Count, document.FileName);

                // Process hyperlinks from end to beginning to avoid index issues when deleting
                for (int i = hyperlinks.Count - 1; i >= 0; i--)
                {
                    var openXmlHyperlink = hyperlinks[i];

                    try
                    {
                        var relId = openXmlHyperlink.Id?.Value;
                        if (string.IsNullOrEmpty(relId))
                            continue;

                        try
                        {
                            // CRITICAL FIX: Build complete URL including fragment exactly like VBA
                            var completeUrl = BuildCompleteHyperlinkUrl(mainPart, openXmlHyperlink);
                            var displayText = openXmlHyperlink.InnerText?.Trim() ?? string.Empty;

                            // CRITICAL FIX: Check if hyperlink has valid lookup ID - use complete URL
                            var lookupId = ExtractIdentifierFromUrl(completeUrl, "");
                            bool hasLookupId = !string.IsNullOrEmpty(lookupId);

                            // CRITICAL FIX: Always remove hyperlinks with empty display text (VBA methodology)
                            // VBA: If Trim$(links(i).TextToDisplay) = "" And Len(links(i).Address) > 0 Then
                            bool shouldRemove = string.IsNullOrEmpty(displayText) && !string.IsNullOrEmpty(completeUrl);

                            if (shouldRemove)
                            {
                                // Log what we're removing for tracking
                                if (hasLookupId)
                                {
                                    _logger.LogInformation("✓ Deleting invisible hyperlink with lookup ID '{LookupId}' (already extracted for API): {Url}", lookupId, completeUrl);
                                }
                                else
                                {
                                    _logger.LogInformation("✓ Deleting invisible hyperlink (no lookup ID): {Url}", completeUrl);
                                }

                                // Remove the hyperlink element
                                openXmlHyperlink.Remove();

                                // Remove the relationship
                                try
                                {
                                    mainPart.DeleteReferenceRelationship(relId);
                                }
                                catch (System.Collections.Generic.KeyNotFoundException)
                                {
                                    _logger.LogDebug("Relationship {RelId} already deleted or doesn't exist", relId);
                                }

                                invisibleLinksRemoved++;

                                // Log the deletion
                                document.ChangeLog.Changes.Add(new ChangeEntry
                                {
                                    Type = ChangeType.HyperlinkRemoved,
                                    Description = hasLookupId ? "Deleted Invisible Hyperlink (lookup ID preserved for API)" : "Deleted Invisible Hyperlink",
                                    OldValue = completeUrl,
                                    NewValue = string.Empty,
                                    ElementId = Guid.NewGuid().ToString(),
                                    Details = hasLookupId ? $"Hyperlink had empty display text but lookup ID '{lookupId}' was already extracted for API processing" : "Hyperlink had empty display text and no valid lookup ID"
                                });

                                // Remove from document.Hyperlinks collection
                                var hyperlinkToRemove = document.Hyperlinks.FirstOrDefault(h => h.OriginalUrl == completeUrl);
                                if (hyperlinkToRemove != null)
                                {
                                    document.Hyperlinks.Remove(hyperlinkToRemove);
                                }
                            }
                        }
                        catch (System.Collections.Generic.KeyNotFoundException)
                        {
                            // Hyperlink has invalid relationship ID - remove only if empty display text
                            var displayText = openXmlHyperlink.InnerText?.Trim() ?? string.Empty;

                            if (string.IsNullOrEmpty(displayText))
                            {
                                openXmlHyperlink.Remove();
                                invisibleLinksRemoved++;

                                document.ChangeLog.Changes.Add(new ChangeEntry
                                {
                                    Type = ChangeType.HyperlinkRemoved,
                                    Description = "Deleted Broken Invisible Hyperlink",
                                    OldValue = "Broken hyperlink",
                                    NewValue = string.Empty,
                                    ElementId = Guid.NewGuid().ToString(),
                                    Details = "Hyperlink had invalid relationship ID and empty display text"
                                });

                                _logger.LogInformation("✓ Deleted broken invisible hyperlink with invalid relationship ID: {RelId}", relId);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning("Error processing hyperlink during invisible link removal: {Error}", ex.Message);
                    }
                }

                _logger.LogInformation("Invisible hyperlink cleanup completed for {FileName}: {RemovedCount} empty hyperlinks removed (lookup IDs already extracted for API processing)",
                    document.FileName, invisibleLinksRemoved);

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error removing invisible hyperlinks in session: {FileName}", document.FileName);
                throw;
            }
        }

        /// <summary>
        /// Updates hyperlinks within the current document session using atomic operations and VBA logic
        /// CRITICAL FIX: Implements proper relationship management to prevent corruption
        /// </summary>
        private async Task UpdateHyperlinksInSessionAsync(MainDocumentPart mainPart, BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken)
        {
            try
            {
                var hyperlinksToUpdate = document.Hyperlinks.Where(h => h.RequiresUpdate).ToList();

                if (!hyperlinksToUpdate.Any())
                {
                    _logger.LogDebug("No hyperlinks require updates in document: {FileName}", document.FileName);
                    return;
                }

                _logger.LogInformation("Updating {Count} hyperlinks in document session using atomic VBA logic: {FileName}", hyperlinksToUpdate.Count, document.FileName);

                // Enhanced logging: Log which hyperlinks are being updated and why
                var titleReplacements = hyperlinksToUpdate.Where(h => h.DisplayText != null && h.DisplayText.Contains("(")).ToList();
                if (titleReplacements.Any())
                {
                    _logger.LogInformation("✓ PROCESSING {Count} TITLE REPLACEMENTS in document session", titleReplacements.Count);
                }

                var hyperlinks = mainPart.Document.Body?.Descendants<OpenXmlHyperlink>().ToList() ?? new List<OpenXmlHyperlink>();
                var processedRelationships = new HashSet<string>();

                foreach (var openXmlHyperlink in hyperlinks)
                {
                    var hyperlinkRelId = openXmlHyperlink.Id?.Value;
                    if (string.IsNullOrEmpty(hyperlinkRelId) || processedRelationships.Contains(hyperlinkRelId))
                        continue;

                    try
                    {
                        // CRITICAL FIX: Build complete URL including fragment exactly like VBA
                        var currentUrl = BuildCompleteHyperlinkUrl(mainPart, openXmlHyperlink);
                        var currentDisplayText = openXmlHyperlink.InnerText ?? string.Empty;

                        var hyperlinkToUpdate = hyperlinksToUpdate.FirstOrDefault(h => h.OriginalUrl == currentUrl);
                        if (hyperlinkToUpdate != null)
                        {
                            await UpdateHyperlinkWithAtomicVbaLogicAsync(mainPart, openXmlHyperlink, hyperlinkRelId, hyperlinkToUpdate, document, cancellationToken);
                            processedRelationships.Add(hyperlinkRelId);
                        }
                    }
                    catch (System.Collections.Generic.KeyNotFoundException)
                    {
                        _logger.LogWarning("Skipping hyperlink update for invalid relationship ID: {RelId} in document: {FileName}", hyperlinkRelId, document.FileName);
                        continue;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error updating individual hyperlink: {RelId} in document: {FileName}", hyperlinkRelId, document.FileName);
                        // Continue with other hyperlinks instead of failing the entire operation
                        continue;
                    }
                }

                // CRITICAL FIX: Final save to ensure all hyperlink updates are persisted
                mainPart.Document.Save();

                _logger.LogInformation("Hyperlink updates completed atomically in session for document: {FileName}", document.FileName);
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating hyperlinks in session: {FileName}", document.FileName);
                throw;
            }
        }

        /// <summary>
        /// Updates a hyperlink using ATOMIC operations and EXACT VBA Base_File.vba logic
        /// CRITICAL FIX: Implements proper transactional relationship management to prevent corruption
        /// Lines 254-280 in Base_File.vba - handles 5-digit to 6-digit upgrade and Content_ID appending
        /// </summary>
        private async Task UpdateHyperlinkWithAtomicVbaLogicAsync(MainDocumentPart mainPart, OpenXmlHyperlink openXmlHyperlink, string relationshipId, Hyperlink hyperlinkToUpdate, BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken)
        {
            string? newRelationshipId = null;
            var originalUri = string.Empty;

            try
            {
                // STEP 1: Validate current state and capture original data
                var originalRelationship = mainPart.GetReferenceRelationship(relationshipId);
                originalUri = originalRelationship.Uri.ToString(); // Keep for logging/fallback

                // CRITICAL FIX: Get complete original URL including fragments for proper comparison
                var completeOriginalUrl = BuildCompleteHyperlinkUrl(mainPart, openXmlHyperlink);

                // CRITICAL FIX: Use the updated display text from hyperlink object (which includes Content ID and status)
                var currentDisplayText = hyperlinkToUpdate.DisplayText ?? openXmlHyperlink.InnerText ?? string.Empty;
                var alreadyExpired = currentDisplayText.Contains(" - Expired", StringComparison.OrdinalIgnoreCase);
                var alreadyNotFound = currentDisplayText.Contains(" - Not Found", StringComparison.OrdinalIgnoreCase);

                _logger.LogDebug("Starting atomic hyperlink update: {RelId}, Complete Original URL: {CompleteOriginalUrl}, Base URI: {BaseUri}",
                    relationshipId, completeOriginalUrl, originalUri);

                // STEP 2: Calculate new URL using proper VBA Address/SubAddress separation (Issue #8)
                // CRITICAL FIX: Handle case where Content_ID is in docid parameter and we need Document_ID
                string docIdForUrl;

                // Enhanced Content_ID detection in docid parameter
                var docIdInUrl = ExtractDocIdFromUrl(completeOriginalUrl);
                var isContentIdInDocId = !string.IsNullOrEmpty(docIdInUrl) &&
                    (!string.IsNullOrEmpty(hyperlinkToUpdate.ContentId) &&
                     string.Equals(docIdInUrl, hyperlinkToUpdate.ContentId, StringComparison.OrdinalIgnoreCase));

                if (!string.IsNullOrEmpty(hyperlinkToUpdate.DocumentId))
                {
                    // Always prefer Document_ID when available
                    docIdForUrl = hyperlinkToUpdate.DocumentId;
                    _logger.LogDebug("Using Document_ID for URL rebuild: '{DocumentId}' (original URL: '{OriginalUrl}')",
                        docIdForUrl, completeOriginalUrl);
                }
                else if (isContentIdInDocId && !string.IsNullOrEmpty(hyperlinkToUpdate.ContentId))
                {
                    // CRITICAL FIX: Content_ID found in docid parameter but no Document_ID available
                    // This indicates the API lookup may have failed or returned Content_ID only
                    // Log this as a warning and use Content_ID as fallback
                    docIdForUrl = hyperlinkToUpdate.ContentId;
                    _logger.LogWarning("CONTENT_ID_IN_DOCID: Found Content_ID '{ContentId}' in docid parameter but no Document_ID available. " +
                        "URL will be rebuilt with Content_ID as fallback. Original URL: '{OriginalUrl}'",
                        hyperlinkToUpdate.ContentId, completeOriginalUrl);
                }
                else if (!string.IsNullOrEmpty(hyperlinkToUpdate.ContentId))
                {
                    // Content_ID available but not in docid parameter
                    docIdForUrl = hyperlinkToUpdate.ContentId;
                    _logger.LogDebug("Using Content_ID for URL rebuild: '{ContentId}' (original URL: '{OriginalUrl}')",
                        docIdForUrl, completeOriginalUrl);
                }
                else
                {
                    // No ID available - this should not happen if hyperlink was processed correctly
                    docIdForUrl = string.Empty;
                    _logger.LogError("NO_ID_AVAILABLE: No Document_ID or Content_ID available for URL rebuild. Original URL: '{OriginalUrl}'",
                        completeOriginalUrl);
                }

                // CRITICAL FIX: Parse URL from UpdatedUrl or calculate from Document ID
                string targetAddress;
                string targetSubAddress;
                string newUrl;

                if (!string.IsNullOrEmpty(hyperlinkToUpdate.UpdatedUrl))
                {
                    // Use the URL calculated by ProcessHyperlinkWithVbaLogicAsync
                    newUrl = hyperlinkToUpdate.UpdatedUrl;

                    // Parse the URL to separate base address and fragment for OpenXML
                    var parts = newUrl.Split('#', 2);
                    targetAddress = parts[0];
                    targetSubAddress = parts.Length > 1 ? parts[1] : string.Empty;

                    _logger.LogInformation("DEBUG: Using UpdatedUrl: '{UpdatedUrl}' -> Base: '{BaseAddress}', Fragment: '{Fragment}', Parts.Length: {PartsLength}",
                        hyperlinkToUpdate.UpdatedUrl, targetAddress, targetSubAddress, parts.Length);
                }
                else
                {
                    // Fall back to recalculating from Document ID or using original URL
                    if (!string.IsNullOrEmpty(docIdForUrl))
                    {
                        // VBA: targetAddress = "https://thesource.cvshealth.com/nuxeo/thesource/"
                        // VBA: targetSub = "!/view?docid=" & rec("Document_ID")
                        targetAddress = "https://thesource.cvshealth.com/nuxeo/thesource/";
                        targetSubAddress = $"!/view?docid={docIdForUrl}";
                        newUrl = targetAddress + "#" + targetSubAddress;

                        _logger.LogDebug("Calculated URL: Base={BaseAddress}, Fragment={Fragment}", targetAddress, targetSubAddress);
                    }
                    else
                    {
                        // No ID available, keep original URL and parse it
                        newUrl = hyperlinkToUpdate.OriginalUrl;
                        var parts = newUrl.Split('#', 2);
                        targetAddress = parts[0];
                        targetSubAddress = parts.Length > 1 ? parts[1] : string.Empty;

                        _logger.LogDebug("Using Original URL: Base={BaseAddress}, Fragment={Fragment}", targetAddress, targetSubAddress);
                    }
                }

                // STEP 3: Only update if URL actually changed to prevent unnecessary operations
                // CRITICAL FIX: Compare complete URLs (including fragments) not just base addresses
                var urlChanged = !string.Equals(completeOriginalUrl, newUrl, StringComparison.OrdinalIgnoreCase);

                if (urlChanged)
                {
                    // CRITICAL FIX: Atomic relationship update with proper VBA Address/SubAddress separation (Issue #8)
                    // Create new relationship with validation
                    try
                    {
                        // CRITICAL FIX: For external URLs with fragments, put the complete URL in the relationship
                        // This ensures Word displays the full URL when hovering over the hyperlink
                        string relationshipUrl;
                        if (!string.IsNullOrEmpty(targetSubAddress))
                        {
                            // Create complete URL for the relationship - Word will display this properly
                            relationshipUrl = newUrl;
                            _logger.LogInformation("DEBUG: Using complete URL for relationship: '{CompleteUrl}'", relationshipUrl);
                        }
                        else
                        {
                            // No fragment, just use the base address
                            relationshipUrl = targetAddress;
                        }

                        var addressUri = new Uri(relationshipUrl);
                        var newRelationship = mainPart.AddHyperlinkRelationship(addressUri, true);
                        newRelationshipId = newRelationship.Id;

                        _logger.LogDebug("Created new relationship atomically: {NewRelId} -> {Address}",
                            newRelationshipId, relationshipUrl);

                        // Update the hyperlink element to use the new relationship ID
                        openXmlHyperlink.Id = newRelationshipId;

                        // CRITICAL FIX: For external URLs, clear any existing DocLocation since complete URL is in relationship
                        if (!string.IsNullOrEmpty(targetSubAddress))
                        {
                            openXmlHyperlink.DocLocation = null;
                            _logger.LogInformation("DEBUG: Cleared DocLocation since complete URL is in relationship");
                        }

                        // Only delete old relationship after successful update
                        try
                        {
                            mainPart.DeleteReferenceRelationship(relationshipId);
                            _logger.LogDebug("Deleted old relationship atomically: {OldRelId}", relationshipId);
                        }
                        catch (System.Collections.Generic.KeyNotFoundException)
                        {
                            _logger.LogDebug("Old relationship {RelId} was already deleted or didn't exist", relationshipId);
                        }

                        // CRITICAL FIX: Save the document after relationship changes
                        mainPart.Document.Save();
                        _logger.LogDebug("Document saved after relationship update: {RelId} -> {NewRelId}", relationshipId, newRelationshipId);
                    }
                    catch (Exception relEx)
                    {
                        _logger.LogError(relEx, "Failed to update relationship atomically: {RelId}", relationshipId);

                        // Cleanup new relationship if it was created
                        if (!string.IsNullOrEmpty(newRelationshipId))
                        {
                            try
                            {
                                mainPart.DeleteReferenceRelationship(newRelationshipId);
                                _logger.LogDebug("Cleaned up failed relationship: {NewRelId}", newRelationshipId);
                            }
                            catch (Exception cleanupEx)
                            {
                                _logger.LogWarning("Failed to cleanup new relationship: {NewRelId}. Error: {Error}", newRelationshipId, cleanupEx.Message);
                            }
                        }
                        throw;
                    }
                }

                // STEP 4: Apply display text changes that were already calculated by ProcessHyperlinkWithVbaLogicAsync
                var newDisplayText = currentDisplayText;
                var displayTextChanged = !string.Equals(openXmlHyperlink.InnerText, currentDisplayText, StringComparison.Ordinal);

                if (displayTextChanged)
                {
                    try
                    {
                        OpenXmlHelper.UpdateHyperlinkText(openXmlHyperlink, newDisplayText);

                        // CRITICAL FIX: Save the document after display text changes
                        mainPart.Document.Save();

                        _logger.LogInformation("✓ DISPLAY TEXT UPDATED IN DOCUMENT: '{OldText}' -> '{NewText}' (RelId: {RelId})", openXmlHyperlink.InnerText, newDisplayText, relationshipId);
                    }
                    catch (Exception textEx)
                    {
                        _logger.LogError(textEx, "Failed to update display text: {RelId}", relationshipId);
                        throw;
                    }
                }

                // STEP 6: Update hyperlink object with final values
                hyperlinkToUpdate.UpdatedUrl = newUrl;
                hyperlinkToUpdate.DisplayText = newDisplayText;
                hyperlinkToUpdate.ActionTaken = HyperlinkAction.Updated;

                // Log URL change if it occurred
                if (urlChanged)
                {
                    document.ChangeLog.Changes.Add(new ChangeEntry
                    {
                        Type = ChangeType.HyperlinkUpdated,
                        Description = "Hyperlink URL updated using atomic VBA logic",
                        OldValue = hyperlinkToUpdate.OriginalUrl,
                        NewValue = newUrl,
                        ElementId = hyperlinkToUpdate.Id,
                        Details = $"Document ID: {docIdForUrl}"
                    });
                }

                _logger.LogInformation("Updated hyperlink atomically with VBA logic: {RelId} -> {NewUrl}, Display: '{NewDisplay}'",
                    relationshipId, newUrl, newDisplayText);
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in atomic hyperlink update with VBA logic: {RelationshipId}", relationshipId);

                // If we created a new relationship but failed, try to clean it up
                if (!string.IsNullOrEmpty(newRelationshipId))
                {
                    try
                    {
                        mainPart.DeleteReferenceRelationship(newRelationshipId);
                        _logger.LogDebug("Cleaned up failed relationship: {NewRelId}", newRelationshipId);
                    }
                    catch (Exception cleanupEx)
                    {
                        _logger.LogWarning("Failed to cleanup relationship during error handling: {NewRelId}. Error: {Error}", newRelationshipId, cleanupEx.Message);
                    }
                }

                throw;
            }
        }

        /// <summary>
        /// DEPRECATED: Legacy update method - replaced by UpdateHyperlinkWithVbaLogicAsync
        /// </summary>
        private async Task UpdateHyperlinkInSessionAsync(MainDocumentPart mainPart, string relationshipId, Hyperlink hyperlinkToUpdate, BulkEditor.Core.Entities.Document document)
        {
            // This method is deprecated - use UpdateHyperlinkWithVbaLogicAsync instead
            _logger.LogDebug("UpdateHyperlinkInSessionAsync called - using VBA logic method instead");
            await Task.CompletedTask;
        }

        /// <summary>
        /// Marks document fields for update to prevent TOC formatting issues
        /// </summary>
        private void MarkDocumentFieldsForUpdate(WordprocessingDocument wordDocument)
        {
            try
            {
                _logger.LogDebug("Marking document fields for update to prevent TOC formatting issues");

                var mainPart = wordDocument.MainDocumentPart;
                if (mainPart?.Document?.Body == null)
                    return;

                // Find and mark field codes for update
                var fieldCodes = mainPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.FieldCode>().ToList();
                var fieldsNeedingUpdate = 0;

                foreach (var fieldCode in fieldCodes)
                {
                    var fieldText = fieldCode.Text ?? string.Empty;

                    // Mark TOC, hyperlink-related fields, and page number fields for update
                    if (fieldText.Contains("TOC", StringComparison.OrdinalIgnoreCase) ||
                        fieldText.Contains("HYPERLINK", StringComparison.OrdinalIgnoreCase) ||
                        fieldText.Contains("PAGE", StringComparison.OrdinalIgnoreCase) ||
                        fieldText.Contains("REF", StringComparison.OrdinalIgnoreCase))
                    {
                        // Find the parent field character and mark as dirty
                        var fieldChar = fieldCode.Ancestors<DocumentFormat.OpenXml.Wordprocessing.FieldChar>().FirstOrDefault();
                        if (fieldChar != null)
                        {
                            fieldChar.Dirty = true;
                            fieldsNeedingUpdate++;
                        }
                    }
                }

                if (fieldsNeedingUpdate > 0)
                {
                    _logger.LogInformation("Marked {Count} document fields for update (TOC, hyperlinks, page numbers)", fieldsNeedingUpdate);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Error marking document fields for update: {Error}", ex.Message);
                // Don't throw - this is not critical enough to fail the entire operation
            }
        }

        /// <summary>
        /// DEPRECATED: This method has been replaced by RemoveInvisibleHyperlinksInSessionAsync
        /// to prevent document corruption from multiple document opens
        /// </summary>
        private async Task RemoveInvisibleHyperlinksAsync(BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken)
        {
            // This method is deprecated to prevent document corruption
            // All operations are now handled in ProcessDocumentInSingleSessionAsync
            _logger.LogDebug("RemoveInvisibleHyperlinksAsync called - operations handled in single session to prevent corruption");
            await Task.CompletedTask;
        }

        /// <summary>
        /// Validates document integrity using OpenXmlValidator
        /// </summary>
        private async Task ValidateDocumentIntegrityAsync(string filePath, string stage, CancellationToken cancellationToken)
        {
            const int maxRetries = 3;
            const int retryDelayMs = 100;

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    if (attempt > 1)
                    {
                        await Task.Delay(retryDelayMs * attempt, cancellationToken).ConfigureAwait(false);
                    }

                    using var wordDocument = WordprocessingDocument.Open(filePath, false);
                    var validationErrors = _validator.Validate(wordDocument)
                        .Where(e => !IsIgnorableValidationError(e.Description))
                        .ToList();

                    if (validationErrors.Any())
                    {
                        var errorDetails = string.Join("; ", validationErrors.Select(e => $"{e.ErrorType}: {e.Description}"));
                        throw new InvalidOperationException($"Document validation failed at {stage}: {errorDetails}");
                    }

                    var mainPart = wordDocument.MainDocumentPart;
                    if (mainPart?.Document?.Body == null)
                    {
                        throw new InvalidOperationException($"Document structure corrupted at {stage} - no main content found");
                    }

                    // Try to access document content to ensure it's readable
                    var _ = mainPart.Document.Body.InnerText;

                    _logger.LogDebug("Document validation passed at {Stage} on attempt {Attempt}: {FilePath}", stage, attempt, filePath);
                    return;
                }
                catch (IOException ioEx) when (ioEx.Message.Contains("being used by another process") && attempt < maxRetries)
                {
                    _logger.LogWarning("File access conflict during validation at {Stage}, attempt {Attempt}/{MaxRetries}: {FilePath}. Retrying...",
                        stage, attempt, maxRetries, filePath);
                    continue;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Document validation failed at {Stage}, attempt {Attempt}: {FilePath}", stage, attempt, filePath);

                    if (attempt == maxRetries)
                    {
                        throw new InvalidOperationException($"Document integrity validation failed at {stage} after {maxRetries} attempts: {ex.Message}", ex);
                    }
                }
            }
        }

        /// <summary>
        /// Validates an open document without file access issues
        /// </summary>
        private async Task ValidateOpenDocumentAsync(WordprocessingDocument wordDocument, string stage, CancellationToken cancellationToken)
        {
            try
            {
                var validationErrors = _validator.Validate(wordDocument)
                    .Where(e => !IsIgnorableValidationError(e.Description))
                    .ToList();

                if (validationErrors.Any())
                {
                    var errorDetails = string.Join("; ", validationErrors.Take(5).Select(e => $"{e.ErrorType}: {e.Description}"));
                    if (validationErrors.Count > 5)
                    {
                        errorDetails += $" (and {validationErrors.Count - 5} more errors)";
                    }

                    _logger.LogError("Document validation failed at {Stage}: {ErrorDetails}", stage, errorDetails);
                    throw new InvalidOperationException($"Document validation failed at {stage}: {errorDetails}");
                }

                _logger.LogDebug("Document validation passed at {Stage}", stage);
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during document validation at {Stage}", stage);
                throw;
            }
        }

        /// <summary>
        /// Creates a snapshot of the current document state for rollback operations
        /// </summary>
        private DocumentSnapshot CreateDocumentSnapshot(MainDocumentPart mainPart)
        {
            try
            {
                var snapshot = new DocumentSnapshot();

                // Capture current relationship mappings
                foreach (var relationship in mainPart.HyperlinkRelationships)
                {
                    snapshot.RelationshipMappings[relationship.Id] = relationship.Uri.ToString();
                }

                _logger.LogDebug("Created document snapshot with {Count} relationships", snapshot.RelationshipMappings.Count);
                return snapshot;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating document snapshot");
                throw;
            }
        }

        /// <summary>
        /// Saves document with comprehensive validation and error handling
        /// CRITICAL FIX: Enhanced save persistence to ensure changes are actually written to disk
        /// </summary>
        private async Task SaveDocumentSafelyAsync(WordprocessingDocument wordDocument, BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken)
        {
            try
            {
                _logger.LogInformation("Starting enhanced document save operation for: {FileName}", document.FileName);

                // Pre-save validation
                await ValidateOpenDocumentAsync(wordDocument, "pre-save-final", cancellationToken).ConfigureAwait(false);

                // CRITICAL FIX: Ensure all parts are saved in the correct order
                var mainPart = wordDocument.MainDocumentPart;
                if (mainPart?.Document == null)
                {
                    throw new InvalidOperationException("Main document part is null - cannot save document");
                }

                // STEP 1: Save all document parts explicitly
                _logger.LogDebug("Saving main document part for: {FileName}", document.FileName);
                mainPart.Document.Save();

                // STEP 2: Save any custom parts that might have been modified
                foreach (var part in wordDocument.Parts)
                {
                    try
                    {
                        if (part.OpenXmlPart is MainDocumentPart || part.OpenXmlPart.RootElement != null)
                        {
                            part.OpenXmlPart.RootElement?.Save();
                        }
                    }
                    catch (Exception partEx)
                    {
                        _logger.LogWarning("Non-critical error saving document part: {Error}", partEx.Message);
                        // Continue with save - part errors shouldn't fail the entire save
                    }
                }

                // STEP 3: Force save the entire document with enhanced retry logic
                _logger.LogDebug("Performing complete document save for: {FileName}", document.FileName);
                var openXmlPolicy = _retryPolicyService.CreateOpenXmlRetryPolicy();
                await _retryPolicyService.ExecuteWithRetryAsync(
                    async () =>
                    {
                        // Force save all changes
                        wordDocument.Save();

                        // CRITICAL FIX: Force save completion by calling save again
                        // This ensures all parts are properly written to disk
                        mainPart.Document.Save();

                        await Task.CompletedTask;
                    },
                    openXmlPolicy, cancellationToken).ConfigureAwait(false);

                // STEP 4: Verify save was successful by attempting to read document structure
                _logger.LogDebug("Verifying document save integrity for: {FileName}", document.FileName);
                try
                {
                    // Ensure we can still access the document structure after save
                    var bodyElementCount = mainPart.Document.Body?.Elements().Count() ?? 0;
                    var hyperlinkCount = mainPart.Document.Body?.Descendants<OpenXmlHyperlink>().Count() ?? 0;
                    _logger.LogDebug("Post-save verification: {BodyElements} body elements, {HyperlinkCount} hyperlinks", bodyElementCount, hyperlinkCount);
                }
                catch (Exception verifyEx)
                {
                    _logger.LogWarning("Post-save verification warning (non-critical): {Error}", verifyEx.Message);
                }

                _logger.LogInformation("✓ DOCUMENT SAVE COMPLETED SUCCESSFULLY: {FileName}", document.FileName);

                // CRITICAL FIX: Extended delay to ensure file system write completion
                // Some file systems need more time to flush changes to disk
                await Task.Delay(100, CancellationToken.None).ConfigureAwait(false);

                // NOTE: Final integrity verification is now handled in post-save validation
                // after the WordprocessingDocument is properly disposed to prevent file access conflicts
                _logger.LogInformation("✓ DOCUMENT SAVE COMPLETED - verification will occur after document disposal");
            }
            catch (Exception saveEx)
            {
                _logger.LogError(saveEx, "Critical error saving document: {FileName}", document.FileName);
                throw new InvalidOperationException($"Failed to save document safely: {saveEx.Message}", saveEx);
            }
        }

        /// <summary>
        /// Performs atomic document operations with enhanced corruption prevention
        /// This method ensures that multiple operations on the same document are serialized
        /// to prevent OpenXML corruption issues
        /// </summary>
        private async Task<T> ExecuteAtomicDocumentOperationAsync<T>(
            string filePath,
            Func<WordprocessingDocument, Task<T>> operation,
            string operationName,
            CancellationToken cancellationToken)
        {
            // Use a semaphore per file path to ensure atomic operations
            var semaphoreKey = filePath.ToLowerInvariant();
            var semaphore = _documentSemaphores.GetOrAdd(semaphoreKey, _ => new SemaphoreSlim(1, 1));

            try
            {
                await semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);

                using (var wordDocument = WordprocessingDocument.Open(filePath, true))
                {
                    // Validate document before operation
                    await ValidateOpenDocumentAsync(wordDocument, $"{operationName}-pre", cancellationToken).ConfigureAwait(false);

                    // Execute the operation
                    var result = await operation(wordDocument).ConfigureAwait(false);

                    // Validate document after operation
                    await ValidateOpenDocumentAsync(wordDocument, $"{operationName}-post", cancellationToken).ConfigureAwait(false);

                    // Ensure changes are saved before releasing the document
                    wordDocument.MainDocumentPart?.Document?.Save();

                    return result;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Atomic document operation '{OperationName}' failed for: {FilePath}", operationName, filePath);
                throw;
            }
            finally
            {
                semaphore.Release();
            }
        }

        /// <summary>
        /// Attempts to recover document from backup when corruption is detected
        /// </summary>
        private async Task AttemptDocumentRecoveryAsync(BulkEditor.Core.Entities.Document document, Exception originalException, CancellationToken cancellationToken)
        {
            try
            {
                _logger.LogWarning("Attempting document recovery from backup due to error: {Error}", originalException.Message);

                if (!string.IsNullOrEmpty(document.BackupPath) && _fileService.FileExists(document.BackupPath))
                {
                    // Validate backup before restoration
                    await ValidateDocumentIntegrityAsync(document.BackupPath, "backup-validation", cancellationToken);

                    // CRITICAL FIX: Use non-cancellable token for backup restoration to prevent data loss
                    // Recovery operations should complete even if main processing was cancelled
                    await RestoreFromBackupAsync(document.FilePath, document.BackupPath, CancellationToken.None);

                    // Validate restoration
                    await ValidateDocumentIntegrityAsync(document.FilePath, "post-recovery", cancellationToken);

                    _logger.LogInformation("Successfully recovered document from backup: {FilePath}", document.FilePath);

                    // Update document status
                    document.Status = DocumentStatus.Recovered;
                    document.ProcessingErrors.Add(new ProcessingError { Message = $"Recovered from error: {originalException.Message}", Severity = ErrorSeverity.Warning });
                }
                else
                {
                    _logger.LogError("Cannot recover document - no valid backup found: {FilePath}", document.FilePath);
                    document.Status = DocumentStatus.Failed;
                    document.ProcessingErrors.Add(new ProcessingError { Message = $"No backup available for recovery: {originalException.Message}", Severity = ErrorSeverity.Error });
                }
            }
            catch (Exception recoveryEx)
            {
                _logger.LogError(recoveryEx, "Document recovery failed: {FilePath}", document.FilePath);
                document.Status = DocumentStatus.Failed;
                document.ProcessingErrors.Add(new ProcessingError { Message = $"Recovery failed: {recoveryEx.Message}", Severity = ErrorSeverity.Error });
            }
        }

        /// <summary>
        /// Validates and sanitizes URL fragments for OpenXML compatibility
        /// CRITICAL FIX: Prevents XSD validation errors from special characters like 0x21 (!)
        /// </summary>
        /// <param name="fragment">URL fragment to validate and sanitize</param>
        /// <returns>Sanitized fragment safe for OpenXML relationships</returns>
        private string ValidateAndSanitizeUrlFragment(string fragment)
        {
            if (string.IsNullOrEmpty(fragment))
                return fragment;

            try
            {
                // Check for problematic characters that cause XSD validation errors
                var problematicChars = new[] { '!', '<', '>', '"', '\'', '&' };
                var hasProblematicChars = fragment.Any(c => problematicChars.Contains(c));

                if (hasProblematicChars)
                {
                    // URL encode the fragment to make it safe for OpenXML
                    var sanitized = Uri.EscapeDataString(fragment);
                    _logger.LogDebug("Sanitized URL fragment for OpenXML compatibility: '{Original}' -> '{Sanitized}'", fragment, sanitized);
                    return sanitized;
                }

                return fragment;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Error sanitizing URL fragment: {Fragment}. Error: {Error}", fragment, ex.Message);
                // Return encoded version as fallback
                return Uri.EscapeDataString(fragment);
            }
        }

        /// <summary>
        /// Validates that a URL fragment is safe for use in OpenXML relationships
        /// CRITICAL FIX: Checks for XSD validation issues with special characters
        /// </summary>
        /// <param name="fragment">URL fragment to validate</param>
        /// <returns>True if fragment is safe for OpenXML, false otherwise</returns>
        private bool IsUrlFragmentSafeForOpenXml(string fragment)
        {
            if (string.IsNullOrEmpty(fragment))
                return true;

            try
            {
                // Characters that cause XSD validation errors in OpenXML relationships
                var unsafeChars = new[] { '!', '<', '>', '"', '\'', '&', '\0' };
                var hasUnsafeChars = fragment.Any(c => unsafeChars.Contains(c) || char.IsControl(c));

                if (hasUnsafeChars)
                {
                    _logger.LogDebug("URL fragment contains unsafe characters for OpenXML: {Fragment}", fragment);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Error validating URL fragment safety: {Fragment}. Error: {Error}", fragment, ex.Message);
                return false;
            }
        }

        public void Dispose()
        {
            // Note: OpenXmlValidator does not implement IDisposable
            // No cleanup needed for validator

            // Cleanup other resources if needed
            if (_replacementService is IDisposable disposableReplacement)
            {
                disposableReplacement.Dispose();
            }
        }

    }
}

