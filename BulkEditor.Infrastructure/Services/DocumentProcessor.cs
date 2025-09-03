using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

                // Extract document metadata
                progress?.Report("Extracting metadata...");
                document.Metadata = ExtractDocumentMetadata(filePath);

                // Extract hyperlinks
                progress?.Report("Extracting hyperlinks...");
                document.Hyperlinks = ExtractHyperlinks(filePath);

                // Validate hyperlinks
                if (document.Hyperlinks.Any())
                {
                    progress?.Report("Validating hyperlinks...");
                    await ValidateHyperlinksAsync(document, cancellationToken);
                }

                // Update hyperlinks if needed
                progress?.Report("Updating hyperlinks...");
                await UpdateHyperlinksAsync(document, cancellationToken);

                // Process replacements (hyperlinks and text)
                progress?.Report("Processing replacements...");
                await _replacementService.ProcessReplacementsAsync(document, cancellationToken);

                // Optimize text if enabled
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

                _logger.LogDebug("Validating {Count} hyperlinks in document: {FileName}", document.Hyperlinks.Count, document.FileName);

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
                        hyperlink.ErrorMessage = result.ErrorMessage;
                        hyperlink.RequiresUpdate = result.RequiresUpdate;
                        hyperlink.LastChecked = DateTime.UtcNow;

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
            try
            {
                var hyperlinksToUpdate = document.Hyperlinks.Where(h => h.RequiresUpdate).ToList();

                if (!hyperlinksToUpdate.Any())
                {
                    _logger.LogDebug("No hyperlinks require updates in document: {FileName}", document.FileName);
                    return document;
                }

                _logger.LogInformation("Updating {Count} hyperlinks in document: {FileName}", hyperlinksToUpdate.Count, document.FileName);

                using var wordDocument = WordprocessingDocument.Open(document.FilePath, true);
                var mainPart = wordDocument.MainDocumentPart;

                if (mainPart?.Document?.Body != null)
                {
                    var hyperlinks = mainPart.Document.Body.Descendants<OpenXmlHyperlink>().ToList();

                    foreach (var openXmlHyperlink in hyperlinks)
                    {
                        var hyperlinkRelId = openXmlHyperlink.Id?.Value;
                        if (string.IsNullOrEmpty(hyperlinkRelId))
                            continue;

                        var relationship = mainPart.GetReferenceRelationship(hyperlinkRelId);
                        var currentUrl = relationship.Uri.ToString();

                        var hyperlinkToUpdate = hyperlinksToUpdate.FirstOrDefault(h => h.OriginalUrl == currentUrl);
                        if (hyperlinkToUpdate != null)
                        {
                            await UpdateHyperlinkInDocument(mainPart, hyperlinkRelId, hyperlinkToUpdate, document);
                        }
                    }

                    // Save the document
                    mainPart.Document.Save();
                }

                await Task.Delay(100, cancellationToken);

                _logger.LogInformation("Hyperlink updates completed for document: {FileName}", document.FileName);
                return document;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating hyperlinks in document: {FileName}", document.FileName);
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
        /// Updates hyperlink title in the Word document
        /// </summary>
        private async Task UpdateHyperlinkTitleInDocumentAsync(string filePath, Hyperlink hyperlink, string newDisplayText, CancellationToken cancellationToken)
        {
            try
            {
                using var wordDocument = WordprocessingDocument.Open(filePath, true);
                var mainPart = wordDocument.MainDocumentPart;

                if (mainPart?.Document?.Body != null)
                {
                    var hyperlinks = mainPart.Document.Body.Descendants<OpenXmlHyperlink>().ToList();

                    foreach (var openXmlHyperlink in hyperlinks)
                    {
                        // Find the matching hyperlink by comparing current display text
                        var currentText = openXmlHyperlink.InnerText;
                        if (currentText == hyperlink.DisplayText)
                        {
                            // Update the display text
                            openXmlHyperlink.RemoveAllChildren();
                            openXmlHyperlink.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(newDisplayText));

                            // Update the hyperlink object
                            hyperlink.DisplayText = newDisplayText;
                            break;
                        }
                    }

                    // Save the document
                    mainPart.Document.Save();
                }

                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating hyperlink title in document: {FilePath}", filePath);
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

                            var relationship = mainPart.GetReferenceRelationship(relId);
                            var url = relationship.Uri.ToString();
                            var displayText = openXmlHyperlink.InnerText;

                            var hyperlink = new Hyperlink
                            {
                                OriginalUrl = url,
                                DisplayText = displayText,
                                LookupId = _hyperlinkValidator.ExtractLookupId(url)
                            };

                            hyperlinks.Add(hyperlink);
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

        private async Task UpdateHyperlinkInDocument(MainDocumentPart mainPart, string relationshipId, Hyperlink hyperlinkToUpdate, BulkEditor.Core.Entities.Document document)
        {
            try
            {
                // Update the relationship target
                var relationship = mainPart.GetReferenceRelationship(relationshipId);

                // Create new URL based on content ID if available
                var newUrl = !string.IsNullOrEmpty(hyperlinkToUpdate.ContentId)
                    ? $"https://example.com/content/{hyperlinkToUpdate.ContentId}"
                    : hyperlinkToUpdate.OriginalUrl;

                // Delete old relationship and create new one
                mainPart.DeleteReferenceRelationship(relationshipId);
                var newRelationship = mainPart.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                    new Uri(newUrl), relationshipId);

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
            var contentIdAdded = changes.Count(c => c.Type == ChangeType.ContentIdAdded);
            var titleReplacements = changes.Count(c => c.Type == ChangeType.TitleReplaced);
            var titleChanges = changes.Count(c => c.Type == ChangeType.PossibleTitleChange);
            var errors = changes.Count(c => c.Type == ChangeType.Error);

            var summaryParts = new List<string>();

            if (hyperlinkUpdates > 0)
                summaryParts.Add($"{hyperlinkUpdates} hyperlinks updated");

            if (contentIdAdded > 0)
                summaryParts.Add($"{contentIdAdded} content IDs added");

            if (titleReplacements > 0)
                summaryParts.Add($"{titleReplacements} titles replaced");

            if (titleChanges > 0)
                summaryParts.Add($"{titleChanges} possible title changes");

            if (errors > 0)
                summaryParts.Add($"{errors} errors");

            document.ChangeLog.Summary = summary + (summaryParts.Any() ? string.Join(", ", summaryParts) : "no changes required");
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