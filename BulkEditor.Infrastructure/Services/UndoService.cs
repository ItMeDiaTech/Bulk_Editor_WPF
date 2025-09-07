using BulkEditor.Core.Interfaces;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace BulkEditor.Infrastructure.Services;

/// <summary>
/// Provides services for undoing file processing operations.
/// </summary>
public class UndoService : IUndoService
{
    private readonly ISessionManager _sessionManager;
    private readonly IBackupService _backupService;
    private readonly ILoggingService _logger;

    /// <summary>
    /// Initializes a new instance of the <see cref="UndoService"/> class.
    /// </summary>
    /// <param name="sessionManager">The session manager.</param>
    /// <param name="backupService">The backup service.</param>
    /// <param name="logger">The logging service.</param>
    public UndoService(ISessionManager sessionManager, IBackupService backupService, ILoggingService logger)
    {
        _sessionManager = sessionManager;
        _backupService = backupService;
        _logger = logger;
    }

    /// <inheritdoc />
    public bool CanUndo()
    {
        var currentSession = _sessionManager.GetCurrentSession();
        return currentSession != null && currentSession.ProcessedFiles.Any();
    }

    /// <inheritdoc />
    public async Task<bool> UndoLastSessionAsync()
    {
        var session = _sessionManager.GetCurrentSession();
        if (session == null || !session.ProcessedFiles.Any())
        {
            _logger.LogWarning("Undo operation requested, but no active session or processed files were found.");
            return false;
        }

        _logger.LogInformation("Starting undo for session: {SessionId}", session.SessionId);
        bool allSucceeded = true;

        foreach (var (originalPath, backupPath) in session.ProcessedFiles)
        {
            try
            {
                await _backupService.RestoreBackupAsync(originalPath, backupPath);
            }
            catch (Exception ex)
            {
                allSucceeded = false;
                _logger.LogError(ex, "Failed to restore file {OriginalPath} from {BackupPath}", originalPath, backupPath);
                // Continue to attempt to restore other files
            }
        }

        if (allSucceeded)
        {
            _logger.LogInformation("Successfully completed undo for session: {SessionId}", session.SessionId);
            // End the session to clean up the used backups
            _sessionManager.EndSession();
        }
        else
        {
            _logger.LogWarning("Undo for session {SessionId} completed with one or more failures.", session.SessionId);
        }

        return allSucceeded;
    }
}