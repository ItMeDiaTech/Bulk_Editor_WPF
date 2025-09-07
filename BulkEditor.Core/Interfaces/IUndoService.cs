using System.Threading.Tasks;

namespace BulkEditor.Core.Interfaces;

/// <summary>
/// Defines a contract for undoing file processing operations.
/// </summary>
public interface IUndoService
{
    /// <summary>
    /// Undoes the last processing session, restoring all modified files from their backups.
    /// </summary>
    /// <returns>A task that represents the asynchronous undo operation. The task result contains a boolean value indicating whether the operation was successful.</returns>
    Task<bool> UndoLastSessionAsync();

    /// <summary>
    /// Checks if there is a session available to be undone.
    /// </summary>
    /// <returns><c>true</c> if an undo operation can be performed; otherwise, <c>false</c>.</returns>
    bool CanUndo();
}