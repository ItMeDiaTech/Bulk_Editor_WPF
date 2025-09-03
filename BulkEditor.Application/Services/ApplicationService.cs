using BulkEditor.Application.Services;
using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Application.Services
{
    /// <summary>
    /// Main application service implementation
    /// </summary>
    public class ApplicationService : IApplicationService
    {
        private readonly IDocumentProcessor _documentProcessor;
        private readonly IFileService _fileService;
        private readonly ILoggingService _logger;

        public ApplicationService(
            IDocumentProcessor documentProcessor,
            IFileService fileService,
            ILoggingService logger)
        {
            _documentProcessor = documentProcessor ?? throw new ArgumentNullException(nameof(documentProcessor));
            _fileService = fileService ?? throw new ArgumentNullException(nameof(fileService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public async Task<Document> ProcessSingleDocumentAsync(string filePath, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
        {
            try
            {
                _logger.LogInformation("Starting single document processing: {FilePath}", filePath);

                progress?.Report($"Validating file: {Path.GetFileName(filePath)}");

                // Validate file first
                var validation = await ValidateFilesAsync(new[] { filePath });
                if (!validation.IsValid)
                {
                    throw new InvalidOperationException($"File validation failed: {string.Join(", ", validation.ErrorMessages)}");
                }

                // Process the document
                var result = await _documentProcessor.ProcessDocumentAsync(filePath, progress, cancellationToken);

                _logger.LogInformation("Single document processing completed: {FilePath}, Status: {Status}",
                    filePath, result.Status);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing single document: {FilePath}", filePath);
                throw;
            }
        }

        public async Task<IEnumerable<Document>> ProcessDocumentsBatchAsync(IEnumerable<string> filePaths, IProgress<BulkEditor.Core.Interfaces.BatchProcessingProgress>? progress = null, CancellationToken cancellationToken = default)
        {
            var stopwatch = Stopwatch.StartNew();
            var filePathsList = filePaths.ToList();

            try
            {
                _logger.LogInformation("Starting batch processing of {Count} files", filePathsList.Count);

                // Validate all files first
                var validation = await ValidateFilesAsync(filePathsList);
                if (validation.InvalidFiles.Any())
                {
                    _logger.LogWarning("Found {Count} invalid files in batch", validation.InvalidFiles.Count);

                    // Report validation issues but continue with valid files
                    var batchProgress = new BulkEditor.Core.Interfaces.BatchProcessingProgress
                    {
                        TotalDocuments = validation.ValidFiles.Count,
                        ProcessedDocuments = 0,
                        FailedDocuments = validation.InvalidFiles.Count,
                        CurrentDocument = "Validation completed"
                    };
                    progress?.Report(batchProgress);
                }

                // Process only valid files
                var results = await _documentProcessor.ProcessDocumentsBatchAsync(validation.ValidFiles, progress, cancellationToken);

                stopwatch.Stop();
                _logger.LogInformation("Batch processing completed in {ElapsedTime}: {TotalFiles} files, {SuccessfulFiles} successful",
                    stopwatch.Elapsed, filePathsList.Count, results.Count(r => r.Status == DocumentStatus.Completed));

                return results;
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                _logger.LogError(ex, "Error in batch processing after {ElapsedTime}", stopwatch.Elapsed);
                throw;
            }
        }

        public async Task<ValidationResult> ValidateFilesAsync(IEnumerable<string> filePaths)
        {
            var result = new ValidationResult();

            try
            {
                foreach (var filePath in filePaths)
                {
                    if (string.IsNullOrWhiteSpace(filePath))
                    {
                        result.ErrorMessages.Add("Empty file path provided");
                        continue;
                    }

                    if (!_fileService.FileExists(filePath))
                    {
                        result.InvalidFiles.Add(filePath);
                        result.ErrorMessages.Add($"File not found: {filePath}");
                        continue;
                    }

                    if (!_fileService.IsValidWordDocument(filePath))
                    {
                        result.InvalidFiles.Add(filePath);
                        result.ErrorMessages.Add($"Not a valid Word document: {filePath}");
                        continue;
                    }

                    // Check if file is accessible (not locked)
                    try
                    {
                        var fileInfo = _fileService.GetFileInfo(filePath);
                        if (fileInfo.IsReadOnly)
                        {
                            result.ErrorMessages.Add($"File is read-only: {filePath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        result.InvalidFiles.Add(filePath);
                        result.ErrorMessages.Add($"Cannot access file: {filePath} - {ex.Message}");
                        continue;
                    }

                    result.ValidFiles.Add(filePath);
                }

                result.IsValid = result.ValidFiles.Any() && !result.InvalidFiles.Any();

                _logger.LogInformation("File validation completed: {ValidCount} valid, {InvalidCount} invalid",
                    result.ValidFiles.Count, result.InvalidFiles.Count);

                return await Task.FromResult(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during file validation");
                result.ErrorMessages.Add($"Validation error: {ex.Message}");
                return result;
            }
        }

        public ProcessingStatistics GetProcessingStatistics(IEnumerable<Document> documents)
        {
            try
            {
                var documentsList = documents.ToList();
                var stats = new ProcessingStatistics
                {
                    TotalDocuments = documentsList.Count,
                    SuccessfulDocuments = documentsList.Count(d => d.Status == DocumentStatus.Completed),
                    FailedDocuments = documentsList.Count(d => d.Status == DocumentStatus.Failed),
                    TotalHyperlinks = documentsList.Sum(d => d.Hyperlinks.Count),
                    UpdatedHyperlinks = documentsList.Sum(d => d.Hyperlinks.Count(h => h.ActionTaken == HyperlinkAction.Updated)),
                    ExpiredHyperlinks = documentsList.Sum(d => d.Hyperlinks.Count(h => h.Status == HyperlinkStatus.Expired)),
                    InvalidHyperlinks = documentsList.Sum(d => d.Hyperlinks.Count(h => h.Status == HyperlinkStatus.Invalid || h.Status == HyperlinkStatus.NotFound))
                };

                // Calculate total processing time
                var processedDocuments = documentsList.Where(d => d.ProcessedAt.HasValue).ToList();
                if (processedDocuments.Any())
                {
                    var minStart = processedDocuments.Min(d => d.CreatedAt);
                    var maxEnd = processedDocuments.Max(d => d.ProcessedAt!.Value);
                    stats.TotalProcessingTime = maxEnd - minStart;
                }

                _logger.LogInformation("Processing statistics calculated: {SuccessfulCount}/{TotalCount} successful",
                    stats.SuccessfulDocuments, stats.TotalDocuments);

                return stats;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error calculating processing statistics");
                return new ProcessingStatistics();
            }
        }

        public async Task<bool> ExportResultsAsync(IEnumerable<Document> documents, string outputPath, ExportFormat format, CancellationToken cancellationToken = default)
        {
            try
            {
                _logger.LogInformation("Exporting results to {OutputPath} in {Format} format", outputPath, format);

                var documentsList = documents.ToList();

                switch (format)
                {
                    case ExportFormat.Json:
                        await ExportToJsonAsync(documentsList, outputPath, cancellationToken);
                        break;
                    case ExportFormat.Csv:
                        await ExportToCsvAsync(documentsList, outputPath, cancellationToken);
                        break;
                    case ExportFormat.Excel:
                        // TODO: Implement Excel export when needed
                        throw new NotImplementedException("Excel export not yet implemented");
                    case ExportFormat.Xml:
                        // TODO: Implement XML export when needed
                        throw new NotImplementedException("XML export not yet implemented");
                    default:
                        throw new ArgumentException($"Unsupported export format: {format}");
                }

                _logger.LogInformation("Export completed successfully: {OutputPath}", outputPath);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting results to {OutputPath}", outputPath);
                return false;
            }
        }

        private async Task ExportToJsonAsync(List<Document> documents, string outputPath, CancellationToken cancellationToken)
        {
            var json = System.Text.Json.JsonSerializer.Serialize(documents, new System.Text.Json.JsonSerializerOptions
            {
                WriteIndented = true
            });

            await File.WriteAllTextAsync(outputPath, json, cancellationToken);
        }

        private async Task ExportToCsvAsync(List<Document> documents, string outputPath, CancellationToken cancellationToken)
        {
            var csv = new System.Text.StringBuilder();
            csv.AppendLine("Document,Status,FilePath,ProcessedAt,HyperlinkCount,UpdatedHyperlinks,ErrorCount");

            foreach (var doc in documents)
            {
                csv.AppendLine($"\"{doc.FileName}\",{doc.Status},\"{doc.FilePath}\",{doc.ProcessedAt},{doc.Hyperlinks.Count},{doc.Hyperlinks.Count(h => h.ActionTaken == HyperlinkAction.Updated)},{doc.ProcessingErrors.Count}");
            }

            await File.WriteAllTextAsync(outputPath, csv.ToString(), cancellationToken);
        }
    }
}