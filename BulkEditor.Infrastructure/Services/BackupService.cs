using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Models;
using BulkEditor.Core.Configuration;
using Microsoft.Extensions.Options;
using System;
using System.IO;
using System.Threading.Tasks;

namespace BulkEditor.Infrastructure.Services;

/// <summary>
/// Provides services for creating and restoring file backups.
/// </summary>
public class BackupService : IBackupService
{
    private readonly ILoggingService _logger;
    private readonly string _backupRoot;

    /// <summary>
    /// Initializes a new instance of the <see cref="BackupService"/> class.
    /// </summary>
    /// <param name="logger">The logging service.</param>
    /// <param name="appSettings">The application settings.</param>
    public BackupService(ILoggingService logger, IOptions<AppSettings> appSettings)
    {
        _logger = logger;
        
        // Get backup directory from settings, with fallback to default
        var backupDir = appSettings.Value.Backup.BackupDirectory;
        
        // If relative path, make it relative to LocalApplicationData
        if (!Path.IsPathRooted(backupDir))
        {
            _backupRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BulkEditor", backupDir);
        }
        else
        {
            _backupRoot = backupDir;
        }
        
        _logger.LogInformation("BackupService initialized with backup directory: {BackupDirectory}", _backupRoot);
    }

    /// <inheritdoc />
    public async Task<string> CreateBackupAsync(string filePath, Session session)
    {
        if (!File.Exists(filePath))
        {
            _logger.LogError("File not found: {FilePath}", filePath);
            throw new FileNotFoundException("The specified file to back up was not found.", filePath);
        }

        // Remove any existing backups for this file before creating a new one
        RemoveExistingBackups(filePath);

        var sessionPath = Path.Combine(_backupRoot, session.SessionId.ToString());
        var backupPath = Path.Combine(sessionPath, Path.GetFileName(filePath));

        try
        {
            // Ensure the session directory exists
            Directory.CreateDirectory(sessionPath);
            await using var sourceStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, true);
            await using var destinationStream = new FileStream(backupPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, true);
            await sourceStream.CopyToAsync(destinationStream);
            _logger.LogInformation("Successfully created backup for {FilePath} at {BackupPath}", filePath, backupPath);
            return backupPath;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to create backup for file: {FilePath}", filePath);
            throw; // Re-throw to allow the caller to handle the failure
        }
    }

    /// <inheritdoc />
    public async Task RestoreBackupAsync(string originalPath, string backupPath)
    {
        if (!File.Exists(backupPath))
        {
            _logger.LogError("Backup file not found: {BackupPath}", backupPath);
            throw new FileNotFoundException("The specified backup file was not found.", backupPath);
        }

        try
        {
            await using var sourceStream = new FileStream(backupPath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, true);
            await using var destinationStream = new FileStream(originalPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, true);
            await sourceStream.CopyToAsync(destinationStream);
            _logger.LogInformation("Successfully restored {OriginalPath} from {BackupPath}", originalPath, backupPath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to restore file {OriginalPath} from backup: {BackupPath}", originalPath, backupPath);
            throw;
        }
    }

    /// <inheritdoc />
    public void ClearBackups(Session session)
    {
        var sessionPath = Path.Combine(_backupRoot, session.SessionId.ToString());
        try
        {
            if (Directory.Exists(sessionPath))
            {
                Directory.Delete(sessionPath, true);
                _logger.LogInformation("Cleared backup directory for session: {SessionId}", session.SessionId);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to clear backups for session: {SessionId}", session.SessionId);
        }
    }

    /// <inheritdoc />
    public bool HasExistingBackup(string filePath)
    {
        if (!File.Exists(filePath))
        {
            return false;
        }

        var fileName = Path.GetFileName(filePath);
        
        try
        {
            if (!Directory.Exists(_backupRoot))
            {
                return false;
            }

            // Check all session directories for a backup of this file
            var sessionDirectories = Directory.GetDirectories(_backupRoot);
            foreach (var sessionDir in sessionDirectories)
            {
                var backupPath = Path.Combine(sessionDir, fileName);
                if (File.Exists(backupPath))
                {
                    _logger.LogDebug("Found existing backup for {FileName} in session {SessionDir}", fileName, Path.GetFileName(sessionDir));
                    return true;
                }
            }

            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to check for existing backup of file: {FilePath}", filePath);
            return false;
        }
    }

    /// <inheritdoc />
    public void RemoveExistingBackups(string filePath)
    {
        if (!File.Exists(filePath))
        {
            _logger.LogWarning("Cannot remove backups for non-existent file: {FilePath}", filePath);
            return;
        }

        var fileName = Path.GetFileName(filePath);
        var removedCount = 0;
        
        try
        {
            if (!Directory.Exists(_backupRoot))
            {
                return;
            }

            // Find and remove all existing backups for this file across all sessions
            var sessionDirectories = Directory.GetDirectories(_backupRoot);
            foreach (var sessionDir in sessionDirectories)
            {
                var backupPath = Path.Combine(sessionDir, fileName);
                if (File.Exists(backupPath))
                {
                    File.Delete(backupPath);
                    removedCount++;
                    _logger.LogInformation("Removed existing backup: {BackupPath}", backupPath);

                    // Clean up empty session directories
                    if (!Directory.GetFiles(sessionDir).Any() && !Directory.GetDirectories(sessionDir).Any())
                    {
                        Directory.Delete(sessionDir);
                        _logger.LogInformation("Removed empty session directory: {SessionDir}", sessionDir);
                    }
                }
            }

            if (removedCount > 0)
            {
                _logger.LogInformation("Removed {Count} existing backup(s) for file: {FileName}", removedCount, fileName);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to remove existing backups for file: {FilePath}", filePath);
            throw;
        }
    }
}