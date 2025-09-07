using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Models;
using System;
using System.IO;

namespace BulkEditor.Infrastructure.Services;

/// <summary>
/// Manages processing sessions for the application.
/// </summary>
public class SessionManager : ISessionManager
{
    private readonly ILoggingService _logger;
    private Session? _currentSession;
    private readonly string _backupRoot;

    /// <summary>
    /// Initializes a new instance of the <see cref="SessionManager"/> class.
    /// </summary>
    /// <param name="logger">The logging service.</param>
    public SessionManager(ILoggingService logger)
    {
        _logger = logger;
        _backupRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BulkEditor", "backups");
        Directory.CreateDirectory(_backupRoot);
    }

    /// <inheritdoc />
    public Session StartSession()
    {
        // End the previous session to perform cleanup
        EndSession();

        _currentSession = new Session();
        _logger.LogInformation("Started new session: {SessionId}", _currentSession.SessionId);

        var sessionPath = Path.Combine(_backupRoot, _currentSession.SessionId.ToString());
        Directory.CreateDirectory(sessionPath);

        return _currentSession;
    }

    /// <inheritdoc />
    public void EndSession()
    {
        if (_currentSession == null)
        {
            return;
        }

        var sessionPath = Path.Combine(_backupRoot, _currentSession.SessionId.ToString());

        try
        {
            if (Directory.Exists(sessionPath))
            {
                Directory.Delete(sessionPath, true);
                _logger.LogInformation("Cleaned up and ended session: {SessionId}", _currentSession.SessionId);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to clean up session directory: {SessionPath}", sessionPath);
        }
        finally
        {
            _currentSession = null;
        }
    }

    /// <inheritdoc />
    public Session? GetCurrentSession()
    {
        return _currentSession;
    }

    /// <inheritdoc />
    public void AddFileToSession(string originalPath, string backupPath)
    {
        if (_currentSession == null)
        {
            _logger.LogWarning("Cannot add file to session. No active session found.");
            return;
        }

        if (!_currentSession.ProcessedFiles.TryAdd(originalPath, backupPath))
        {
            _logger.LogWarning("File '{OriginalPath}' has already been added to the current session.", originalPath);
        }
    }
}