using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Core.Interfaces
{
    /// <summary>
    /// Interface for file system operations
    /// </summary>
    public interface IFileService
    {
        /// <summary>
        /// Checks if a file exists
        /// </summary>
        bool FileExists(string filePath);

        /// <summary>
        /// Checks if a directory exists
        /// </summary>
        bool DirectoryExists(string directoryPath);

        /// <summary>
        /// Creates a directory if it doesn't exist
        /// </summary>
        void CreateDirectory(string directoryPath);

        /// <summary>
        /// Copies a file to a new location
        /// </summary>
        Task CopyFileAsync(string sourcePath, string destinationPath, CancellationToken cancellationToken = default);

        /// <summary>
        /// Moves a file to a new location
        /// </summary>
        Task MoveFileAsync(string sourcePath, string destinationPath, CancellationToken cancellationToken = default);

        /// <summary>
        /// Deletes a file
        /// </summary>
        Task DeleteFileAsync(string filePath, CancellationToken cancellationToken = default);

        /// <summary>
        /// Gets file information
        /// </summary>
        FileInfo GetFileInfo(string filePath);

        /// <summary>
        /// Gets all Word documents in a directory
        /// </summary>
        IEnumerable<string> GetWordDocuments(string directoryPath, bool recursive = false);

        /// <summary>
        /// Gets the size of a file in bytes
        /// </summary>
        long GetFileSize(string filePath);

        /// <summary>
        /// Gets the last modified date of a file
        /// </summary>
        DateTime GetLastModified(string filePath);

        /// <summary>
        /// Creates a backup file with timestamp
        /// </summary>
        Task<string> CreateBackupAsync(string filePath, string backupDirectory, CancellationToken cancellationToken = default);

        /// <summary>
        /// Validates if a file is a valid Word document
        /// </summary>
        bool IsValidWordDocument(string filePath);

        /// <summary>
        /// Gets available disk space in bytes
        /// </summary>
        long GetAvailableDiskSpace(string path);
    }
}