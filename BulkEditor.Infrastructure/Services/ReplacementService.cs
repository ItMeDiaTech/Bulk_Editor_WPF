using BulkEditor.Core.Configuration;
using BulkEditor.Core.Entities;
using BulkEditor.Core.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Main implementation of replacement service that coordinates hyperlink and text replacements
    /// </summary>
    public class ReplacementService : IReplacementService
    {
        private readonly IHyperlinkReplacementService _hyperlinkReplacementService;
        private readonly ITextReplacementService _textReplacementService;
        private readonly ILoggingService _logger;
        private readonly AppSettings _appSettings;

        public ReplacementService(
            IHyperlinkReplacementService hyperlinkReplacementService,
            ITextReplacementService textReplacementService,
            ILoggingService logger,
            AppSettings appSettings)
        {
            _hyperlinkReplacementService = hyperlinkReplacementService ?? throw new ArgumentNullException(nameof(hyperlinkReplacementService));
            _textReplacementService = textReplacementService ?? throw new ArgumentNullException(nameof(textReplacementService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _appSettings = appSettings ?? throw new ArgumentNullException(nameof(appSettings));
        }

        /// <summary>
        /// NEW METHOD: Processes replacements using an already opened WordprocessingDocument to prevent corruption
        /// </summary>
        public async Task<int> ProcessReplacementsInSessionAsync(DocumentFormat.OpenXml.Packaging.WordprocessingDocument wordDocument, Document document, CancellationToken cancellationToken = default)
        {
            try
            {
                var replacementSettings = _appSettings.Replacement;
                var totalReplacements = 0;

                _logger.LogInformation("Starting replacement processing in session for document: {FileName}", document.FileName);

                // Process hyperlink replacements if enabled
                if (replacementSettings.EnableHyperlinkReplacement && replacementSettings.HyperlinkRules.Any())
                {
                    _logger.LogDebug("Processing hyperlink replacements in session for document: {FileName}", document.FileName);
                    totalReplacements += await _hyperlinkReplacementService.ProcessHyperlinkReplacementsInSessionAsync(
                        wordDocument, document, replacementSettings.HyperlinkRules, cancellationToken);
                }

                // Process text replacements if enabled
                if (replacementSettings.EnableTextReplacement && replacementSettings.TextRules.Any())
                {
                    _logger.LogDebug("Processing text replacements in session for document: {FileName}", document.FileName);
                    totalReplacements += await _textReplacementService.ProcessTextReplacementsInSessionAsync(
                        wordDocument, document, replacementSettings.TextRules, cancellationToken);
                }

                _logger.LogInformation("Replacement processing completed in session for document: {FileName}, total replacements: {Count}", document.FileName, totalReplacements);
                return totalReplacements;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing replacements in session for document: {FileName}", document.FileName);

                // Add error to change log
                document.ChangeLog.Changes.Add(new ChangeEntry
                {
                    Type = ChangeType.Error,
                    Description = "Error during replacement processing",
                    Details = ex.Message
                });

                throw;
            }
        }

        /// <summary>
        /// LEGACY METHOD: Opens document independently - can cause corruption
        /// </summary>
        [System.Obsolete("Use ProcessReplacementsInSessionAsync to prevent file corruption")]
        public async Task<Document> ProcessReplacementsAsync(Document document, CancellationToken cancellationToken = default)
        {
            try
            {
                var replacementSettings = _appSettings.Replacement;

                _logger.LogInformation("Starting replacement processing for document: {FileName}", document.FileName);

                // Process hyperlink replacements if enabled
                if (replacementSettings.EnableHyperlinkReplacement && replacementSettings.HyperlinkRules.Any())
                {
                    _logger.LogDebug("Processing hyperlink replacements for document: {FileName}", document.FileName);
                    document = await _hyperlinkReplacementService.ProcessHyperlinkReplacementsAsync(
                        document, replacementSettings.HyperlinkRules, cancellationToken);
                }

                // Process text replacements if enabled
                if (replacementSettings.EnableTextReplacement && replacementSettings.TextRules.Any())
                {
                    _logger.LogDebug("Processing text replacements for document: {FileName}", document.FileName);
                    document = await _textReplacementService.ProcessTextReplacementsAsync(
                        document, replacementSettings.TextRules, cancellationToken);
                }

                _logger.LogInformation("Replacement processing completed for document: {FileName}", document.FileName);
                return document;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing replacements for document: {FileName}", document.FileName);

                // Add error to change log
                document.ChangeLog.Changes.Add(new ChangeEntry
                {
                    Type = ChangeType.Error,
                    Description = "Error during replacement processing",
                    Details = ex.Message
                });

                throw;
            }
        }

        public async Task<ReplacementValidationResult> ValidateReplacementRulesAsync(IEnumerable<object> rules, CancellationToken cancellationToken = default)
        {
            try
            {
                var result = new ReplacementValidationResult { IsValid = true };

                await Task.Run(() =>
                {
                    foreach (var rule in rules)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        switch (rule)
                        {
                            case HyperlinkReplacementRule hyperlinkRule:
                                ValidateHyperlinkRule(hyperlinkRule, result);
                                break;
                            case TextReplacementRule textRule:
                                ValidateTextRule(textRule, result);
                                break;
                            default:
                                result.ValidationErrors.Add($"Unknown rule type: {rule.GetType().Name}");
                                result.InvalidRulesCount++;
                                break;
                        }
                    }
                }, cancellationToken);

                result.IsValid = !result.ValidationErrors.Any();

                _logger.LogInformation("Replacement rule validation completed: {ValidCount} valid, {InvalidCount} invalid",
                    result.ValidRulesCount, result.InvalidRulesCount);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error validating replacement rules");
                throw;
            }
        }

        private void ValidateHyperlinkRule(HyperlinkReplacementRule rule, ReplacementValidationResult result)
        {
            var errors = new List<string>();

            // Validate title to match is not empty
            if (string.IsNullOrWhiteSpace(rule.TitleToMatch))
            {
                errors.Add($"Hyperlink rule {rule.Id}: Title to match cannot be empty");
            }

            // Validate Content ID format (should contain 6 digits)
            if (string.IsNullOrWhiteSpace(rule.ContentId))
            {
                errors.Add($"Hyperlink rule {rule.Id}: Content ID cannot be empty");
            }
            else
            {
                var contentIdRegex = new Regex(@"[0-9]{6}", RegexOptions.Compiled);
                if (!contentIdRegex.IsMatch(rule.ContentId))
                {
                    errors.Add($"Hyperlink rule {rule.Id}: Content ID must contain a 6-digit number");
                }
            }

            if (errors.Any())
            {
                result.ValidationErrors.AddRange(errors);
                result.InvalidRulesCount++;
            }
            else
            {
                result.ValidRulesCount++;
            }
        }

        private void ValidateTextRule(TextReplacementRule rule, ReplacementValidationResult result)
        {
            var errors = new List<string>();

            // Validate source text is not empty
            if (string.IsNullOrWhiteSpace(rule.SourceText))
            {
                errors.Add($"Text rule {rule.Id}: Source text cannot be empty");
            }

            // Validate replacement text is not empty
            if (string.IsNullOrWhiteSpace(rule.ReplacementText))
            {
                errors.Add($"Text rule {rule.Id}: Replacement text cannot be empty");
            }

            // Check for potential infinite loop (source = replacement)
            if (!string.IsNullOrWhiteSpace(rule.SourceText) && !string.IsNullOrWhiteSpace(rule.ReplacementText))
            {
                if (rule.SourceText.Trim().Equals(rule.ReplacementText.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    errors.Add($"Text rule {rule.Id}: Source text and replacement text cannot be the same");
                }
            }

            if (errors.Any())
            {
                result.ValidationErrors.AddRange(errors);
                result.InvalidRulesCount++;
            }
            else
            {
                result.ValidRulesCount++;
            }
        }
    }
}