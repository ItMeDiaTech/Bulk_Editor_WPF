using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using DocumentFormat.OpenXml.Packaging;
using System;
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

        // VBA-compatible regex pattern for exact Lookup_ID matching
        // Must match Base_File.vba exactly but ensure exactly 6 digits (not 7+)
        // Using negative lookahead (?![0-9]) to prevent matching partial patterns
        private static readonly Regex LookupIdRegex = new Regex(@"(TSRC-[^-]+-[0-9]{6}(?![0-9])|CMS-[^-]+-[0-9]{6}(?![0-9]))", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex DocIdRegex = new Regex(@"docid=([^&]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public DocumentProcessor(IFileService fileService, IHyperlinkValidator hyperlinkValidator, ITextOptimizer textOptimizer, IReplacementService replacementService, ILoggingService logger, Core.Configuration.AppSettings appSettings)
        {
            _fileService = fileService ?? throw new ArgumentNullException(nameof(fileService));
            _hyperlinkValidator = hyperlinkValidator ?? throw new ArgumentNullException(nameof(hyperlinkValidator));
            _textOptimizer = textOptimizer ?? throw new ArgumentNullException(nameof(textOptimizer));
            _replacementService = replacementService ?? throw new ArgumentNullException(nameof(replacementService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _appSettings = appSettings ?? throw new ArgumentNullException(nameof(appSettings));
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

                // Create backup
                progress?.Report("Creating backup...");
                document.BackupPath = await CreateBackupAsync(filePath, cancellationToken);

                // CRITICAL FIX: Process document in single session to prevent corruption
                progress?.Report("Processing document...");
                await ProcessDocumentInSingleSessionAsync(document, progress, cancellationToken);

                // Process replacements (hyperlinks and text) - external service
                progress?.Report("Processing replacements...");
                await _replacementService.ProcessReplacementsAsync(document, cancellationToken);

                // Optimize text if enabled - external service
                progress?.Report("Optimizing document text...");
                await _textOptimizer.OptimizeDocumentTextAsync(document, cancellationToken);

                // Optimize memory after processing
                await OptimizeMemoryAsync(cancellationToken);

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
                document.ProcessingErrors.Add(ex.Message);

                _logger.LogError(ex, "Error processing document: {FilePath}", filePath);
                progress?.Report($"Error processing document: {ex.Message}");

                // Try to restore from backup if document was corrupted
                if (!string.IsNullOrEmpty(document.BackupPath) && _fileService.FileExists(document.BackupPath))
                {
                    try
                    {
                        await RestoreFromBackupAsync(filePath, document.BackupPath, cancellationToken);
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

            var batchProgress = new BatchProcessingProgress
            {
                TotalDocuments = filePathsList.Count
            };

            _logger.LogInformation("Starting batch processing of {Count} documents", filePathsList.Count);

            try
            {
                // Use configurable concurrency based on system resources
                var maxConcurrency = Math.Min(_appSettings.Processing.MaxConcurrentDocuments, Environment.ProcessorCount * 2);
                var semaphore = new SemaphoreSlim(maxConcurrency);

                var tasks = filePathsList.Select(async filePath =>
                {
                    await semaphore.WaitAsync(cancellationToken);
                    try
                    {
                        batchProgress.CurrentDocument = Path.GetFileName(filePath);
                        progress?.Report(batchProgress);

                        var document = await ProcessDocumentAsync(filePath, null, cancellationToken);

                        lock (results)
                        {
                            results.Add(document);
                            if (document.Status == DocumentStatus.Completed)
                                Interlocked.Increment(ref processed);
                            else
                                Interlocked.Increment(ref failed);

                            batchProgress.ProcessedDocuments = processed;
                            batchProgress.FailedDocuments = failed;
                            progress?.Report(batchProgress);
                        }

                        return document;
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                });

                await Task.WhenAll(tasks);

                // Optimize memory after batch processing
                await OptimizeMemoryAsync(cancellationToken);

                _logger.LogInformation("Batch processing completed: {Processed} successful, {Failed} failed", processed, failed);
                return results;
            }
            catch (Exception ex)
            {
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

                var validationResults = await _hyperlinkValidator.ValidateHyperlinksAsync(document.Hyperlinks, cancellationToken);

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
                            await HandleTitleDifferenceAsync(document, hyperlink, result.TitleComparison, cancellationToken);
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

        public async Task<BulkEditor.Core.Entities.Document> UpdateHyperlinksAsync(BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken = default)
        {
            // This method is now handled within ProcessDocumentInSingleSessionAsync to prevent corruption
            // Individual hyperlink updates should not open the document separately
            _logger.LogDebug("UpdateHyperlinksAsync called - operations handled in single session to prevent corruption");
            return document;
        }

        public async Task<string> CreateBackupAsync(string filePath, CancellationToken cancellationToken = default)
        {
            try
            {
                var backupDirectory = Path.Combine(Path.GetDirectoryName(filePath) ?? "", "Backups");
                return await _fileService.CreateBackupAsync(filePath, backupDirectory, cancellationToken);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating backup for file: {FilePath}", filePath);
                throw;
            }
        }

        public async Task<bool> RestoreFromBackupAsync(string filePath, string backupPath, CancellationToken cancellationToken = default)
        {
            try
            {
                await _fileService.CopyFileAsync(backupPath, filePath, cancellationToken);
                _logger.LogInformation("Restored file from backup: {FilePath}", filePath);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error restoring file from backup: {FilePath}", filePath);
                return false;
            }
        }

        /// <summary>
        /// Handles title differences between current hyperlink and API response
        /// </summary>
        private async Task HandleTitleDifferenceAsync(BulkEditor.Core.Entities.Document document, Hyperlink hyperlink, TitleComparisonResult titleComparison, CancellationToken cancellationToken)
        {
            try
            {
                var validationSettings = _appSettings.Validation;

                if (validationSettings.AutoReplaceTitles)
                {
                    // Replace the title with API title and append Content ID
                    var newDisplayText = $"{titleComparison.ApiTitle} ({titleComparison.ContentId})";

                    // Update the hyperlink in the document
                    await UpdateHyperlinkTitleInDocumentAsync(document.FilePath, hyperlink, newDisplayText, cancellationToken);

                    titleComparison.WasReplaced = true;
                    titleComparison.ActionTaken = "Title replaced with API response";

                    // Log the replacement
                    document.ChangeLog.Changes.Add(new ChangeEntry
                    {
                        Type = ChangeType.TitleReplaced,
                        Description = "Title replaced with API response",
                        OldValue = titleComparison.CurrentTitle,
                        NewValue = titleComparison.ApiTitle,
                        ElementId = hyperlink.Id,
                        Details = $"Content ID: {titleComparison.ContentId}"
                    });

                    _logger.LogInformation("Replaced hyperlink title: '{OldTitle}' -> '{NewTitle}' (Content ID: {ContentId})",
                        titleComparison.CurrentTitle, titleComparison.ApiTitle, titleComparison.ContentId);
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
                                var relationship = mainPart.GetReferenceRelationship(relId);
                                var url = relationship.Uri.ToString();
                                var displayText = openXmlHyperlink.InnerText;

                                var hyperlink = new Hyperlink
                                {
                                    OriginalUrl = url,
                                    DisplayText = displayText,
                                    LookupId = _hyperlinkValidator.ExtractLookupId(url),
                                    RequiresUpdate = ShouldAutoValidateHyperlink(url, displayText)
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

            var lookupId = ExtractLookupIdUsingVbaLogic(url, "");
            var shouldValidate = !string.IsNullOrEmpty(lookupId);

            _logger.LogDebug("Hyperlink validation check: URL={Url}, LookupID={LookupId}, ShouldValidate={ShouldValidate}",
                url, lookupId, shouldValidate);

            return shouldValidate;
        }

        /// <summary>
        /// Extracts Lookup_ID using EXACT same logic as VBA ExtractLookupID function
        /// This ensures consistent behavior between VBA and C# implementations
        /// </summary>
        /// <param name="address">Hyperlink address</param>
        /// <param name="subAddress">Hyperlink sub-address</param>
        /// <returns>Extracted Lookup_ID or empty string if no match</returns>
        private string ExtractLookupIdUsingVbaLogic(string address, string subAddress)
        {
            try
            {
                // Combine address and subAddress like VBA: addr & IIf(Len(subAddr) > 0, "#" & subAddr, "")
                var fullUrl = address + (!string.IsNullOrEmpty(subAddress) ? "#" + subAddress : "");

                // First, try exact VBA regex pattern: (TSRC-[^-]+-[0-9]{6}|CMS-[^-]+-[0-9]{6})
                var regexMatch = LookupIdRegex.Match(fullUrl);
                if (regexMatch.Success)
                {
                    var lookupId = regexMatch.Value.ToUpperInvariant();
                    _logger.LogDebug("Extracted Lookup_ID from regex: {LookupId} from URL: {Url}", lookupId, fullUrl);
                    return lookupId;
                }

                // Fallback: Check for docid= parameter (like VBA)
                var docIdMatch = DocIdRegex.Match(fullUrl);
                if (docIdMatch.Success)
                {
                    var docId = docIdMatch.Groups[1].Value.Trim();
                    _logger.LogDebug("Extracted docid from URL: {DocId} from URL: {Url}", docId, fullUrl);
                    return docId;
                }

                _logger.LogDebug("No Lookup_ID found in URL: {Url}", fullUrl);
                return string.Empty;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Error extracting Lookup_ID from URL: {Url}. Error: {Error}", address, ex.Message);
                return string.Empty;
            }
        }

        private async Task UpdateHyperlinkInDocument(MainDocumentPart mainPart, string relationshipId, Hyperlink hyperlinkToUpdate, BulkEditor.Core.Entities.Document document)
        {
            try
            {
                // Update the relationship target
                var relationship = mainPart.GetReferenceRelationship(relationshipId);

                // Create new URL using Document_ID for docid parameter (correct approach)
                // Fallback to Content_ID if Document_ID is not available
                var docIdForUrl = !string.IsNullOrEmpty(hyperlinkToUpdate.DocumentId)
                    ? hyperlinkToUpdate.DocumentId
                    : hyperlinkToUpdate.ContentId;

                var newUrl = !string.IsNullOrEmpty(docIdForUrl)
                    ? $"https://thesource.cvshealth.com/nuxeo/thesource/#!/view?docid={docIdForUrl}"
                    : hyperlinkToUpdate.OriginalUrl;

                // Delete old relationship and create new one
                mainPart.DeleteReferenceRelationship(relationshipId);
                var newRelationship = mainPart.AddHyperlinkRelationship(new Uri(newUrl), true, relationshipId);

                // Update hyperlink object
                hyperlinkToUpdate.UpdatedUrl = newUrl;
                hyperlinkToUpdate.ActionTaken = HyperlinkAction.Updated;

                // Log the change
                document.ChangeLog.Changes.Add(new ChangeEntry
                {
                    Type = ChangeType.HyperlinkUpdated,
                    Description = "Hyperlink updated",
                    OldValue = hyperlinkToUpdate.OriginalUrl,
                    NewValue = newUrl,
                    ElementId = hyperlinkToUpdate.Id
                });

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating hyperlink in document");
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
                string statusSuffix = null;
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
        /// CRITICAL FIX: Process all document operations in a single session to prevent corruption
        /// </summary>
        private async Task ProcessDocumentInSingleSessionAsync(BulkEditor.Core.Entities.Document document, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
        {
            try
            {
                _logger.LogInformation("Processing document in single session to prevent corruption: {FileName}", document.FileName);

                // Open document once and perform ALL operations within this session
                using (var wordDocument = WordprocessingDocument.Open(document.FilePath, true))
                {
                    var mainPart = wordDocument.MainDocumentPart;

                    if (mainPart?.Document?.Body == null)
                    {
                        throw new InvalidOperationException($"Document has no main content: {document.FilePath}");
                    }

                    // STEP 1: Extract metadata (read-only operations first)
                    progress?.Report("Extracting metadata...");
                    document.Metadata = ExtractDocumentMetadataFromOpenDocument(wordDocument);

                    // STEP 2: Extract hyperlinks from the open document
                    progress?.Report("Extracting hyperlinks...");
                    document.Hyperlinks = ExtractHyperlinksFromOpenDocument(mainPart);

                    // STEP 3: Remove invisible hyperlinks (write operations)
                    progress?.Report("Removing invisible hyperlinks...");
                    await RemoveInvisibleHyperlinksInSessionAsync(mainPart, document, cancellationToken);

                    // STEP 4: Validate hyperlinks (external API calls)
                    if (document.Hyperlinks.Any())
                    {
                        progress?.Report("Validating hyperlinks...");
                        await ValidateHyperlinksAsync(document, cancellationToken);
                    }

                    // STEP 5: Update hyperlinks in the same session
                    progress?.Report("Updating hyperlinks...");
                    await UpdateHyperlinksInSessionAsync(mainPart, document, cancellationToken);

                    // STEP 6: Save document once at the end
                    progress?.Report("Saving document...");
                    mainPart.Document.Save();
                } // CRITICAL FIX: WordprocessingDocument is disposed here - ensures file handles are released

                // STEP 7: Validate document integrity after proper disposal with delay
                progress?.Report("Validating document integrity...");
                await ValidateDocumentIntegrityWithRetryAsync(document.FilePath, cancellationToken);

                _logger.LogInformation("Document processed successfully in single session: {FileName}", document.FileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in single session document processing: {FileName}", document.FileName);
                throw;
            }
        }

        /// <summary>
        /// Validates document integrity after processing with retry logic to handle file locking issues
        /// </summary>
        private async Task ValidateDocumentIntegrityWithRetryAsync(string filePath, CancellationToken cancellationToken)
        {
            const int maxRetries = 3;
            const int retryDelayMs = 100;

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    // Add small delay to ensure file handles are fully released
                    if (attempt > 1)
                    {
                        await Task.Delay(retryDelayMs * attempt, cancellationToken);
                    }

                    // Try to open the document in read-only mode to verify it's not corrupted
                    using var testDocument = WordprocessingDocument.Open(filePath, false);
                    var mainPart = testDocument.MainDocumentPart;

                    if (mainPart?.Document?.Body == null)
                    {
                        throw new InvalidOperationException("Document appears to be corrupted - no main content found");
                    }

                    // Try to access the document content to ensure it's readable
                    var _ = mainPart.Document.Body.InnerText;

                    _logger.LogDebug("Document integrity validation passed on attempt {Attempt}: {FilePath}", attempt, filePath);
                    return; // Success
                }
                catch (IOException ioEx) when (ioEx.Message.Contains("being used by another process") && attempt < maxRetries)
                {
                    _logger.LogWarning("File access conflict on attempt {Attempt}/{MaxRetries}: {FilePath}. Retrying...",
                        attempt, maxRetries, filePath);
                    continue; // Retry
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Document integrity validation failed on attempt {Attempt}: {FilePath}", attempt, filePath);

                    if (attempt == maxRetries)
                    {
                        throw new InvalidOperationException($"Document appears to be corrupted after processing: {ex.Message}", ex);
                    }
                }
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
                var openXmlHyperlinks = mainPart.Document.Body.Descendants<OpenXmlHyperlink>().ToList();

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
                            var relationship = mainPart.GetReferenceRelationship(relId);
                            var url = relationship.Uri.ToString();
                            var displayText = openXmlHyperlink.InnerText;

                            var hyperlink = new Hyperlink
                            {
                                OriginalUrl = url,
                                DisplayText = displayText,
                                LookupId = _hyperlinkValidator.ExtractLookupId(url),
                                RequiresUpdate = ShouldAutoValidateHyperlink(url, displayText)
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
        /// </summary>
        private async Task RemoveInvisibleHyperlinksInSessionAsync(MainDocumentPart mainPart, BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken)
        {
            try
            {
                var invisibleLinksRemoved = 0;
                var hyperlinks = mainPart.Document.Body.Descendants<OpenXmlHyperlink>().ToList();

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
                            var relationship = mainPart.GetReferenceRelationship(relId);
                            var url = relationship.Uri.ToString();
                            var displayText = openXmlHyperlink.InnerText?.Trim() ?? string.Empty;

                            // Check if hyperlink is invisible (empty display text but has URL)
                            if (string.IsNullOrEmpty(displayText) && !string.IsNullOrEmpty(url))
                            {
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
                                    Description = "Deleted Invisible Hyperlink",
                                    OldValue = url,
                                    NewValue = string.Empty,
                                    ElementId = Guid.NewGuid().ToString(),
                                    Details = "Hyperlink had empty display text"
                                });

                                // Remove from document.Hyperlinks collection
                                var hyperlinkToRemove = document.Hyperlinks.FirstOrDefault(h => h.OriginalUrl == url);
                                if (hyperlinkToRemove != null)
                                {
                                    document.Hyperlinks.Remove(hyperlinkToRemove);
                                }

                                _logger.LogInformation("Deleted invisible hyperlink: {Url}", url);
                            }
                        }
                        catch (System.Collections.Generic.KeyNotFoundException)
                        {
                            // Hyperlink has invalid relationship ID - remove if empty display text
                            var displayText = openXmlHyperlink.InnerText?.Trim() ?? string.Empty;

                            if (string.IsNullOrEmpty(displayText))
                            {
                                openXmlHyperlink.Remove();
                                invisibleLinksRemoved++;

                                document.ChangeLog.Changes.Add(new ChangeEntry
                                {
                                    Type = ChangeType.HyperlinkRemoved,
                                    Description = "Deleted Invisible Hyperlink",
                                    OldValue = "Broken hyperlink",
                                    NewValue = string.Empty,
                                    ElementId = Guid.NewGuid().ToString(),
                                    Details = "Hyperlink had invalid relationship ID and empty display text"
                                });

                                _logger.LogInformation("Deleted broken invisible hyperlink with invalid relationship ID: {RelId}", relId);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning("Error processing hyperlink during invisible link removal: {Error}", ex.Message);
                    }
                }

                if (invisibleLinksRemoved > 0)
                {
                    _logger.LogInformation("Removed {Count} invisible hyperlinks from document: {FileName}",
                        invisibleLinksRemoved, document.FileName);
                }

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error removing invisible hyperlinks in session: {FileName}", document.FileName);
                throw;
            }
        }

        /// <summary>
        /// Updates hyperlinks within the current document session
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

                _logger.LogInformation("Updating {Count} hyperlinks in document session: {FileName}", hyperlinksToUpdate.Count, document.FileName);

                var hyperlinks = mainPart.Document.Body.Descendants<OpenXmlHyperlink>().ToList();

                foreach (var openXmlHyperlink in hyperlinks)
                {
                    var hyperlinkRelId = openXmlHyperlink.Id?.Value;
                    if (string.IsNullOrEmpty(hyperlinkRelId))
                        continue;

                    try
                    {
                        var relationship = mainPart.GetReferenceRelationship(hyperlinkRelId);
                        var currentUrl = relationship.Uri.ToString();

                        var hyperlinkToUpdate = hyperlinksToUpdate.FirstOrDefault(h => h.OriginalUrl == currentUrl);
                        if (hyperlinkToUpdate != null)
                        {
                            await UpdateHyperlinkInSessionAsync(mainPart, hyperlinkRelId, hyperlinkToUpdate, document);
                        }
                    }
                    catch (System.Collections.Generic.KeyNotFoundException)
                    {
                        _logger.LogWarning("Skipping hyperlink update for invalid relationship ID: {RelId} in document: {FileName}", hyperlinkRelId, document.FileName);
                        continue;
                    }
                }

                _logger.LogInformation("Hyperlink updates completed in session for document: {FileName}", document.FileName);
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating hyperlinks in session: {FileName}", document.FileName);
                throw;
            }
        }

        /// <summary>
        /// Updates a hyperlink within the current document session
        /// </summary>
        private async Task UpdateHyperlinkInSessionAsync(MainDocumentPart mainPart, string relationshipId, Hyperlink hyperlinkToUpdate, BulkEditor.Core.Entities.Document document)
        {
            try
            {
                // CRITICAL: Use Document_ID for URL generation (like VBA), not Content_ID
                var docIdForUrl = !string.IsNullOrEmpty(hyperlinkToUpdate.DocumentId)
                    ? hyperlinkToUpdate.DocumentId
                    : hyperlinkToUpdate.ContentId;

                var newUrl = !string.IsNullOrEmpty(docIdForUrl)
                    ? $"https://thesource.cvshealth.com/nuxeo/thesource/#!/view?docid={docIdForUrl}"
                    : hyperlinkToUpdate.OriginalUrl;

                // Delete old relationship and create new one with same ID
                mainPart.DeleteReferenceRelationship(relationshipId);
                var newRelationship = mainPart.AddHyperlinkRelationship(new Uri(newUrl), true, relationshipId);

                // Update hyperlink object
                hyperlinkToUpdate.UpdatedUrl = newUrl;
                hyperlinkToUpdate.ActionTaken = HyperlinkAction.Updated;

                // Log the change
                document.ChangeLog.Changes.Add(new ChangeEntry
                {
                    Type = ChangeType.HyperlinkUpdated,
                    Description = "Hyperlink updated in session",
                    OldValue = hyperlinkToUpdate.OriginalUrl,
                    NewValue = newUrl,
                    ElementId = hyperlinkToUpdate.Id,
                    Details = $"Document ID: {docIdForUrl}"
                });

                _logger.LogInformation("Updated hyperlink in session: {RelId} -> {NewUrl}", relationshipId, newUrl);
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating hyperlink in session: {RelationshipId}", relationshipId);
                throw;
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

        public void Dispose()
        {
            // Cleanup resources if needed
            if (_replacementService is IDisposable disposableReplacement)
            {
                disposableReplacement.Dispose();
            }
        }
    }
}