using BulkEditor.Core.Models;
using System;

namespace BulkEditor.Core.Interfaces;

/// <summary>
/// Defines a contract for managing processing sessions.
/// </summary>
public interface ISessionManager
{
    /// <summary>
    /// Starts a new processing session.
    /// </summary>
    /// <returns>The newly created session.</returns>
    Session StartSession();

    /// <summary>
    /// Ends the current processing session and performs cleanup.
    /// </summary>
    void EndSession();

    /// <summary>
    /// Gets the current active session.
    /// </summary>
    /// <returns>The current session, or null if no session is active.</returns>
    Session? GetCurrentSession();

    /// <summary>
    /// Adds a processed file to the current session.
    /// </summary>
    /// <param name="originalPath">The original path of the file.</param>
    /// <param name="backupPath">The path to the created backup file.</param>
    void AddFileToSession(string originalPath, string backupPath);
}