using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using BulkEditor.Infrastructure.Utilities;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
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

        // CRITICAL FIX: Exact VBA pattern match without negative lookahead (Issue #1)
        // VBA: .Pattern = "(TSRC-[^-]+-[0-9]{6}|CMS-[^-]+-[0-9]{6})"
        // VBA: .IgnoreCase = True
        private static readonly Regex LookupIdRegex = new Regex(@"(TSRC-[^-]+-[0-9]{6}|CMS-[^-]+-[0-9]{6})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex DocIdRegex = new Regex(@"docid=([^&]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // OpenXML validator for document integrity checks
        private readonly OpenXmlValidator _validator = new OpenXmlValidator();

        // Add a HashSet to store ignorable validation error descriptions for performance
        private static readonly HashSet<string> IgnorableValidationErrorDescriptions = new HashSet<string>
        {
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
            "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:tr'."
        };

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

                // NOTE: Replacements and text optimization are now handled within the single session
                // to prevent file corruption from multiple document opens
                progress?.Report("Document processing completed in single session");

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
                document.ProcessingErrors.Add(new ProcessingError { Message = ex.Message, Severity = ErrorSeverity.Error });

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
        /// CRITICAL FIX: Now includes proper docid fallback and URL encoding handling
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

                // CRITICAL FIX: First, try exact VBA regex pattern with case-insensitive matching
                var regexMatch = LookupIdRegex.Match(fullUrl);
                if (regexMatch.Success)
                {
                    var lookupId = regexMatch.Value.ToUpperInvariant();
                    _logger.LogDebug("Extracted Lookup_ID from regex: {LookupId} from URL: {Url}", lookupId, fullUrl);
                    return lookupId;
                }

                // CRITICAL FIX: Fallback docid extraction exactly like VBA (Issue #3)
                // VBA: ElseIf InStr(1, full, "docid=", vbTextCompare) > 0 Then
                if (fullUrl.IndexOf("docid=", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // VBA: ExtractLookupID = Trim$(Split(Split(full, "docid=")(1), "&")(0))
                    var parts = fullUrl.Split(new[] { "docid=" }, StringSplitOptions.None);
                    if (parts.Length > 1)
                    {
                        var docId = parts[1].Split('&')[0].Trim();
                        // CRITICAL FIX: Handle URL encoding (Issue #3)
                        var decodedDocId = Uri.UnescapeDataString(docId);
                        _logger.LogDebug("Extracted docid fallback from URL: {DocId} (decoded: {DecodedDocId}) from URL: {Url}",
                            docId, decodedDocId, fullUrl);
                        return decodedDocId;
                    }
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

                foreach (var hyperlink in document.Hyperlinks)
                {
                    var lookupId = ExtractLookupIdUsingVbaLogic(hyperlink.OriginalUrl, "");
                    if (!string.IsNullOrEmpty(lookupId) && !idDict.ContainsKey(lookupId))
                    {
                        idDict[lookupId] = true;
                        _logger.LogDebug("Added unique Lookup_ID: {LookupId}", lookupId);
                    }
                }

                if (idDict.Count == 0)
                {
                    _logger.LogInformation("No valid Lookup_IDs found in document: {FileName}", document.FileName);
                    return;
                }

                _logger.LogInformation("Found {Count} unique Lookup_IDs for API processing", idDict.Count);

                // VBA STEP 3: Build JSON & POST (lines 87-128)
                var lookupIds = idDict.Keys.ToArray();

                // CRITICAL FIX: Use HyperlinkReplacementService for API processing with VBA methodology
                // First, we need an HttpService instance with HttpClient
                using var httpClient = new System.Net.Http.HttpClient();
                var httpService = new HttpService(httpClient, _logger);
                var hyperlinkService = new HyperlinkReplacementService(httpService, _logger);
                var apiResult = await hyperlinkService.ProcessApiResponseAsync(lookupIds, cancellationToken);

                if (apiResult.HasError)
                {
                    _logger.LogError("API processing failed: {Error}", apiResult.ErrorMessage);
                    return;
                }

                // VBA STEP 5: Update hyperlinks using dictionary lookup (lines 186-318)
                await UpdateHyperlinksUsingVbaDictionaryLogicAsync(document, apiResult, cancellationToken);

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

                // Process each hyperlink exactly like VBA
                foreach (var hyperlink in document.Hyperlinks)
                {
                    var lookupId = ExtractLookupIdUsingVbaLogic(hyperlink.OriginalUrl, "");
                    if (string.IsNullOrEmpty(lookupId))
                        continue;

                    var dispText = hyperlink.DisplayText ?? string.Empty;
                    var alreadyExpired = dispText.Contains(" - Expired", StringComparison.OrdinalIgnoreCase);
                    var alreadyNotFound = dispText.Contains(" - Not Found", StringComparison.OrdinalIgnoreCase);

                    if (recDict.ContainsKey(lookupId))
                    {
                        // VBA: If recDict.Exists(lookupID) Then Set rec = recDict(lookupID)
                        var rec = recDict[lookupId];
                        await ProcessHyperlinkWithVbaLogicAsync(hyperlink, rec, document, alreadyExpired, alreadyNotFound, cancellationToken);
                    }
                    else if (!alreadyNotFound && !alreadyExpired)
                    {
                        // VBA: ElseIf Not alreadyNotFound And Not alreadyExpired Then
                        // VBA: hl.TextToDisplay = hl.TextToDisplay & " - Not Found"
                        hyperlink.DisplayText += " - Not Found";
                        hyperlink.Status = HyperlinkStatus.NotFound;

                        document.ChangeLog.Changes.Add(new ChangeEntry
                        {
                            Type = ChangeType.HyperlinkStatusAdded,
                            Description = "Hyperlink marked as Not Found (VBA methodology)",
                            OldValue = dispText,
                            NewValue = hyperlink.DisplayText,
                            ElementId = hyperlink.Id,
                            Details = $"Lookup_ID: {lookupId}"
                        });

                        _logger.LogInformation("Marked hyperlink as Not Found: {LookupId}", lookupId);
                    }
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
        /// </summary>
        private async Task ProcessHyperlinkWithVbaLogicAsync(Hyperlink hyperlink, DocumentRecord rec, BulkEditor.Core.Entities.Document document, bool alreadyExpired, bool alreadyNotFound, CancellationToken cancellationToken)
        {
            try
            {
                var dispText = hyperlink.DisplayText ?? string.Empty;

                // VBA: targetAddress = "https://thesource.cvshealth.com/nuxeo/thesource/"
                // VBA: targetSub = "!/view?docid=" & rec("Document_ID")
                var targetAddress = "https://thesource.cvshealth.com/nuxeo/thesource/";
                var targetSub = $"!/view?docid={rec.Document_ID}";

                // VBA: changedURL = (hl.Address <> targetAddress) Or (hl.SubAddress <> targetSub)
                var targetUrl = targetAddress + "#" + targetSub;
                var changedURL = !string.Equals(hyperlink.OriginalUrl, targetUrl, StringComparison.OrdinalIgnoreCase);

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
                                appended = true;
                                _logger.LogInformation("Upgraded 5-digit to 6-digit Content_ID: {Old} -> {New}", pattern5, pattern6);
                            }
                        }
                        // VBA: ElseIf InStr(1, dispText, " (" & last6 & ")", vbTextCompare) = 0 Then
                        else if (!dispText.Contains(pattern6, StringComparison.OrdinalIgnoreCase))
                        {
                            hyperlink.DisplayText = dispText.Trim() + pattern6;
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

            try
            {
                _logger.LogInformation("Processing document in single session with enhanced corruption prevention: {FileName}", document.FileName);

                // STEP 1: Pre-processing validation
                progress?.Report("Validating document before processing...");
                await ValidateDocumentIntegrityAsync(document.FilePath, "pre-processing", cancellationToken);

                // Open document once and perform ALL operations within this session
                using (var wordDocument = WordprocessingDocument.Open(document.FilePath, true))
                {
                    var mainPart = wordDocument.MainDocumentPart;

                    if (mainPart?.Document?.Body == null)
                    {
                        throw new InvalidOperationException($"Document has no main content: {document.FilePath}");
                    }

                    // STEP 2: Initial document validation
                    progress?.Report("Validating document structure...");
                    await ValidateOpenDocumentAsync(wordDocument, "initial", cancellationToken);

                    // STEP 3: Create document snapshot for rollback
                    progress?.Report("Creating document snapshot...");
                    snapshot = CreateDocumentSnapshot(mainPart);

                    // STEP 4: Extract metadata (read-only operations first)
                    progress?.Report("Extracting metadata...");
                    document.Metadata = ExtractDocumentMetadataFromOpenDocument(wordDocument);

                    // STEP 5: Extract hyperlinks from the open document
                    progress?.Report("Extracting hyperlinks...");
                    document.Hyperlinks = ExtractHyperlinksFromOpenDocument(mainPart);

                    // STEP 6: Remove invisible hyperlinks (write operations)
                    progress?.Report("Removing invisible hyperlinks...");
                    await RemoveInvisibleHyperlinksInSessionAsync(mainPart, document, cancellationToken);

                    // STEP 7: Validate after invisible hyperlink removal
                    progress?.Report("Validating after invisible hyperlink removal...");
                    await ValidateOpenDocumentAsync(wordDocument, "post-cleanup", cancellationToken);

                    // STEP 8: Process hyperlinks using VBA UpdateHyperlinksFromAPI workflow
                    if (document.Hyperlinks.Any())
                    {
                        progress?.Report("Processing hyperlinks using VBA methodology...");
                        await ProcessHyperlinksUsingVbaWorkflowAsync(document, cancellationToken);
                    }

                    // STEP 9: Apply hyperlink updates in the document session
                    progress?.Report("Applying hyperlink updates to document...");
                    await UpdateHyperlinksInSessionAsync(mainPart, document, cancellationToken);

                    // STEP 10: Validate after hyperlink updates
                    progress?.Report("Validating after hyperlink updates...");
                    await ValidateOpenDocumentAsync(wordDocument, "post-hyperlinks", cancellationToken);

                    // STEP 11: Process replacements in the same session
                    progress?.Report("Processing replacements...");
                    await _replacementService.ProcessReplacementsInSessionAsync(wordDocument, document, cancellationToken);

                    // STEP 12: Validate after replacements
                    progress?.Report("Validating after replacements...");
                    await ValidateOpenDocumentAsync(wordDocument, "post-replacements", cancellationToken);

                    // STEP 13: Optimize text in the same session (only if enabled)
                    if (_appSettings.Processing.OptimizeText)
                    {
                        progress?.Report("Optimizing document text...");
                        await _textOptimizer.OptimizeDocumentTextInSessionAsync(wordDocument, document, cancellationToken);
                    }
                    else
                    {
                        _logger.LogDebug("Text optimization skipped - disabled in settings for document: {FileName}", document.FileName);
                    }

                    // STEP 14: Final validation before save
                    progress?.Report("Final validation before save...");
                    await ValidateOpenDocumentAsync(wordDocument, "pre-save", cancellationToken);

                    // STEP 15: Save document with enhanced error handling and validation
                    progress?.Report("Saving document with validation...");
                    await SaveDocumentSafelyAsync(wordDocument, document, cancellationToken);

                } // CRITICAL: WordprocessingDocument disposed here - ensures file handles are released

                // STEP 16: Post-save validation
                progress?.Report("Validating document after save...");
                await ValidateDocumentIntegrityAsync(document.FilePath, "post-save", cancellationToken);

                _logger.LogInformation("Document processed successfully with comprehensive validation: {FileName}", document.FileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in single session document processing with rollback attempt: {FileName}", document.FileName);

                // Attempt to restore from backup if we have one
                if (!string.IsNullOrEmpty(document.BackupPath))
                {
                    await AttemptDocumentRecoveryAsync(document, ex, cancellationToken);
                }

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

                var hyperlinks = mainPart.Document.Body.Descendants<OpenXmlHyperlink>().ToList();
                var processedRelationships = new HashSet<string>();

                foreach (var openXmlHyperlink in hyperlinks)
                {
                    var hyperlinkRelId = openXmlHyperlink.Id?.Value;
                    if (string.IsNullOrEmpty(hyperlinkRelId) || processedRelationships.Contains(hyperlinkRelId))
                        continue;

                    try
                    {
                        var relationship = mainPart.GetReferenceRelationship(hyperlinkRelId);
                        var currentUrl = relationship.Uri.ToString();
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
                originalUri = originalRelationship.Uri.ToString();
                var currentDisplayText = openXmlHyperlink.InnerText ?? string.Empty;
                var alreadyExpired = currentDisplayText.Contains(" - Expired", StringComparison.OrdinalIgnoreCase);
                var alreadyNotFound = currentDisplayText.Contains(" - Not Found", StringComparison.OrdinalIgnoreCase);

                _logger.LogDebug("Starting atomic hyperlink update: {RelId}, Original URL: {OriginalUrl}", relationshipId, originalUri);

                // STEP 2: Calculate new URL using proper VBA Address/SubAddress separation (Issue #8)
                var docIdForUrl = !string.IsNullOrEmpty(hyperlinkToUpdate.DocumentId)
                    ? hyperlinkToUpdate.DocumentId
                    : hyperlinkToUpdate.ContentId;

                // CRITICAL FIX: Separate Address and SubAddress exactly like VBA (Issue #8)
                // VBA: targetAddress = "https://thesource.cvshealth.com/nuxeo/thesource/"
                // VBA: targetSub = "!/view?docid=" & rec("Document_ID")
                var targetAddress = "https://thesource.cvshealth.com/nuxeo/thesource/";
                var targetSubAddress = $"!/view?docid={docIdForUrl}";

                // Build complete URL for validation/logging only (NOT for relationship creation)
                var newUrl = !string.IsNullOrEmpty(docIdForUrl)
                    ? targetAddress + "#" + targetSubAddress
                    : hyperlinkToUpdate.OriginalUrl;

                // STEP 3: Only update if URL actually changed to prevent unnecessary operations
                var urlChanged = !string.Equals(originalUri, newUrl, StringComparison.OrdinalIgnoreCase);

                if (urlChanged)
                {
                    // CRITICAL FIX: Atomic relationship update with proper VBA Address/SubAddress separation (Issue #8)
                    // Create new relationship with validation
                    try
                    {
                        // CRITICAL FIX: Use separate Address and SubAddress like VBA (Issue #8)
                        // VBA: .Address = targetAddress, .SubAddress = targetSub
                        var addressUri = new Uri(targetAddress);
                        var newRelationship = mainPart.AddHyperlinkRelationship(addressUri, true, targetSubAddress);
                        newRelationshipId = newRelationship.Id;

                        _logger.LogDebug("Created new relationship atomically with VBA Address/SubAddress: {NewRelId} -> {Address}#{SubAddress}",
                            newRelationshipId, targetAddress, targetSubAddress);

                        // Update the hyperlink element to use the new relationship ID
                        openXmlHyperlink.Id = newRelationshipId;

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

                // STEP 4: EXACT VBA LOGIC: Content_ID appending (lines 254-280 in Base_File.vba)
                var newDisplayText = currentDisplayText;
                var displayTextChanged = false;

                if (!alreadyExpired && !alreadyNotFound && !string.IsNullOrEmpty(hyperlinkToUpdate.ContentId))
                {
                    // Get last 6 and last 5 digits like VBA
                    // CRITICAL FIX: Add proper bounds checking to prevent IndexOutOfRangeException (Issue #12)
                    string last6, last5;
                    if (hyperlinkToUpdate.ContentId.Length >= 6)
                    {
                        last6 = hyperlinkToUpdate.ContentId.Substring(hyperlinkToUpdate.ContentId.Length - 6);
                        last5 = last6.Length > 1 ? last6.Substring(1) : last6; // Safe substring
                    }
                    else if (hyperlinkToUpdate.ContentId.Length == 5)
                    {
                        // Pad 5-digit with leading zero like VBA methodology
                        last6 = "0" + hyperlinkToUpdate.ContentId;
                        last5 = hyperlinkToUpdate.ContentId; // Original 5 digits
                    }
                    else
                    {
                        // Handle shorter content IDs safely
                        last6 = hyperlinkToUpdate.ContentId.PadLeft(6, '0');
                        last5 = last6.Length > 1 ? last6.Substring(1) : last6;
                    }

                    var last5Pattern = $" ({last5})";
                    var last6Pattern = $" ({last6})";

                    // CRITICAL FIX: VBA Logic with safe substring operations (Issue #12)
                    if (currentDisplayText.EndsWith(last5Pattern) && !currentDisplayText.EndsWith(last6Pattern))
                    {
                        // Safe substring operation with bounds checking
                        if (currentDisplayText.Length >= last5Pattern.Length)
                        {
                            newDisplayText = currentDisplayText.Substring(0, currentDisplayText.Length - last5Pattern.Length) + last6Pattern;
                            displayTextChanged = true;
                            _logger.LogInformation("Upgraded 5-digit Content_ID to 6-digit: {Old} -> {New}", last5Pattern, last6Pattern);
                        }
                    }
                    // VBA Logic: If Content_ID not already present, append it
                    else if (!currentDisplayText.Contains(last6Pattern, StringComparison.OrdinalIgnoreCase))
                    {
                        newDisplayText = currentDisplayText.Trim() + last6Pattern;
                        displayTextChanged = true;
                        _logger.LogInformation("Appended Content_ID to hyperlink: {ContentId}", last6);
                    }

                    // Update display text in the document if changed
                    if (displayTextChanged)
                    {
                        try
                        {
                            OpenXmlHelper.UpdateHyperlinkText(openXmlHyperlink, newDisplayText);

                            document.ChangeLog.Changes.Add(new ChangeEntry
                            {
                                Type = ChangeType.ContentIdAdded,
                                Description = "Content ID appended using atomic VBA logic",
                                OldValue = currentDisplayText,
                                NewValue = newDisplayText,
                                ElementId = hyperlinkToUpdate.Id,
                                Details = $"Content ID: {last6}"
                            });
                        }
                        catch (Exception textEx)
                        {
                            _logger.LogError(textEx, "Failed to update display text atomically: {RelId}", relationshipId);
                            throw;
                        }
                    }
                }

                // STEP 5: Handle status suffixes like VBA (Expired/Not Found)
                if (hyperlinkToUpdate.Status == HyperlinkStatus.Expired && !alreadyExpired)
                {
                    newDisplayText += " - Expired";
                    displayTextChanged = true;
                }
                else if (hyperlinkToUpdate.Status == HyperlinkStatus.NotFound && !alreadyNotFound && !alreadyExpired)
                {
                    newDisplayText += " - Not Found";
                    displayTextChanged = true;
                }

                // Apply status suffix changes if needed
                if (displayTextChanged && (hyperlinkToUpdate.Status == HyperlinkStatus.Expired || hyperlinkToUpdate.Status == HyperlinkStatus.NotFound))
                {
                    try
                    {
                        OpenXmlHelper.UpdateHyperlinkText(openXmlHyperlink, newDisplayText);

                        var statusType = hyperlinkToUpdate.Status == HyperlinkStatus.Expired ? "Expired" : "Not Found";
                        document.ChangeLog.Changes.Add(new ChangeEntry
                        {
                            Type = ChangeType.HyperlinkStatusAdded,
                            Description = $"Added {statusType} status atomically",
                            OldValue = currentDisplayText,
                            NewValue = newDisplayText,
                            ElementId = hyperlinkToUpdate.Id
                        });
                    }
                    catch (Exception statusEx)
                    {
                        _logger.LogError(statusEx, "Failed to update status suffix atomically: {RelId}", relationshipId);
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
                        .Where(e => !IgnorableValidationErrorDescriptions.Contains(e.Description))
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
                    .Where(e => !IgnorableValidationErrorDescriptions.Contains(e.Description))
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
        /// </summary>
        private async Task SaveDocumentSafelyAsync(WordprocessingDocument wordDocument, BulkEditor.Core.Entities.Document document, CancellationToken cancellationToken)
        {
            try
            {
                // Pre-save validation
                await ValidateOpenDocumentAsync(wordDocument, "pre-save-final", cancellationToken);

                // Ensure all changes are committed before saving
                var mainPart = wordDocument.MainDocumentPart;
                mainPart?.Document?.Save();

                _logger.LogDebug("Document saved successfully with validation: {FileName}", document.FileName);

                // Small delay to ensure file handles are properly released
                await Task.Delay(50, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception saveEx)
            {
                _logger.LogError(saveEx, "Critical error saving document: {FileName}", document.FileName);
                throw new InvalidOperationException($"Failed to save document safely: {saveEx.Message}", saveEx);
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

                    // Restore from backup
                    await RestoreFromBackupAsync(document.FilePath, document.BackupPath, cancellationToken);

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

