using System;
using System.Collections.Generic;

namespace BulkEditor.Core.Models;

/// <summary>
/// Represents a file processing session, holding data for potential undos.
/// </summary>
public class Session
{
    /// <summary>
    /// Gets the unique identifier for the session.
    /// </summary>
    public Guid SessionId { get; }

    /// <summary>
    /// Gets the time the session started.
    /// </summary>
    public DateTime StartTime { get; }

    /// <summary>
    /// Gets a dictionary mapping original file paths to their backup locations.
    /// </summary>
    public Dictionary<string, string> ProcessedFiles { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="Session"/> class.
    /// </summary>
    public Session()
    {
        SessionId = Guid.NewGuid();
        StartTime = DateTime.Now;
        ProcessedFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }
}