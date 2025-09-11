using BulkEditor.Core.Models;
using System.Threading.Tasks;

namespace BulkEditor.Core.Interfaces;

/// <summary>
/// Defines a contract for creating and restoring file backups.
/// </summary>
public interface IBackupService
{
    /// <summary>
    /// Creates a backup of a file in the specified session's directory.
    /// </summary>
    /// <param name="filePath">The path to the file to back up.</param>
    /// <param name="session">The session to associate the backup with.</param>
    /// <returns>The path to the created backup file.</returns>
    Task<string> CreateBackupAsync(string filePath, Session session);

    /// <summary>
    /// Restores a file from its backup.
    /// </summary>
    /// <param name="originalPath">The original path of the file to restore.</param>
    /// <param name="backupPath">The path of the backup file.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    Task RestoreBackupAsync(string originalPath, string backupPath);

    /// <summary>
    /// Clears all backups associated with a specific session.
    /// </summary>
    /// <param name="session">The session whose backups should be cleared.</param>
    void ClearBackups(Session session);

    /// <summary>
    /// Checks if any backup exists for the specified file across all sessions.
    /// </summary>
    /// <param name="filePath">The path to the file to check for existing backups.</param>
    /// <returns>True if a backup exists for the file, otherwise false.</returns>
    bool HasExistingBackup(string filePath);

    /// <summary>
    /// Removes all existing backups for the specified file across all sessions.
    /// </summary>
    /// <param name="filePath">The path to the file whose backups should be removed.</param>
    void RemoveExistingBackups(string filePath);
}