using BulkEditor.Core.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Implementation of file system operations
    /// </summary>
    public class FileService : IFileService
    {
        private readonly ILoggingService _logger;
        private readonly string[] _wordExtensions = { ".docx", ".docm" };

        public FileService(ILoggingService logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public bool FileExists(string filePath)
        {
            try
            {
                return File.Exists(filePath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking if file exists: {FilePath}", filePath);
                return false;
            }
        }

        public bool DirectoryExists(string directoryPath)
        {
            try
            {
                return Directory.Exists(directoryPath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking if directory exists: {DirectoryPath}", directoryPath);
                return false;
            }
        }

        public void CreateDirectory(string directoryPath)
        {
            try
            {
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                    _logger.LogInformation("Created directory: {DirectoryPath}", directoryPath);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating directory: {DirectoryPath}", directoryPath);
                throw;
            }
        }

        public async Task CopyFileAsync(string sourcePath, string destinationPath, CancellationToken cancellationToken = default)
        {
            try
            {
                // Ensure destination directory exists
                var destinationDir = Path.GetDirectoryName(destinationPath);
                if (!string.IsNullOrEmpty(destinationDir))
                {
                    CreateDirectory(destinationDir);
                }

                await Task.Run(() => File.Copy(sourcePath, destinationPath, overwrite: true), cancellationToken);
                _logger.LogInformation("Copied file from {Source} to {Destination}", sourcePath, destinationPath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error copying file from {Source} to {Destination}", sourcePath, destinationPath);
                throw;
            }
        }

        public async Task MoveFileAsync(string sourcePath, string destinationPath, CancellationToken cancellationToken = default)
        {
            try
            {
                // Ensure destination directory exists
                var destinationDir = Path.GetDirectoryName(destinationPath);
                if (!string.IsNullOrEmpty(destinationDir))
                {
                    CreateDirectory(destinationDir);
                }

                await Task.Run(() => File.Move(sourcePath, destinationPath), cancellationToken);
                _logger.LogInformation("Moved file from {Source} to {Destination}", sourcePath, destinationPath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error moving file from {Source} to {Destination}", sourcePath, destinationPath);
                throw;
            }
        }

        public async Task DeleteFileAsync(string filePath, CancellationToken cancellationToken = default)
        {
            try
            {
                await Task.Run(() => File.Delete(filePath), cancellationToken);
                _logger.LogInformation("Deleted file: {FilePath}", filePath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting file: {FilePath}", filePath);
                throw;
            }
        }

        public FileInfo GetFileInfo(string filePath)
        {
            try
            {
                return new FileInfo(filePath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting file info: {FilePath}", filePath);
                throw;
            }
        }

        public IEnumerable<string> GetWordDocuments(string directoryPath, bool recursive = false)
        {
            try
            {
                var searchOption = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                var files = new List<string>();

                foreach (var extension in _wordExtensions)
                {
                    var pattern = $"*{extension}";
                    files.AddRange(Directory.GetFiles(directoryPath, pattern, searchOption));
                }

                _logger.LogInformation("Found {Count} Word documents in {Directory}", files.Count, directoryPath);
                return files;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting Word documents from directory: {DirectoryPath}", directoryPath);
                throw;
            }
        }

        public long GetFileSize(string filePath)
        {
            try
            {
                var fileInfo = new FileInfo(filePath);
                return fileInfo.Length;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting file size: {FilePath}", filePath);
                throw;
            }
        }

        public DateTime GetLastModified(string filePath)
        {
            try
            {
                var fileInfo = new FileInfo(filePath);
                return fileInfo.LastWriteTime;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting last modified date: {FilePath}", filePath);
                throw;
            }
        }

        public async Task<string> CreateBackupAsync(string filePath, string backupDirectory, CancellationToken cancellationToken = default)
        {
            try
            {
                CreateDirectory(backupDirectory);

                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var extension = Path.GetExtension(filePath);
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var backupFileName = $"{fileName}_backup_{timestamp}{extension}";
                var backupPath = Path.Combine(backupDirectory, backupFileName);

                await CopyFileAsync(filePath, backupPath, cancellationToken);
                _logger.LogInformation("Created backup: {BackupPath}", backupPath);

                return backupPath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating backup for file: {FilePath}", filePath);
                throw;
            }
        }

        public bool IsValidWordDocument(string filePath)
        {
            try
            {
                if (!FileExists(filePath))
                    return false;

                var extension = Path.GetExtension(filePath).ToLowerInvariant();
                return _wordExtensions.Contains(extension);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error validating Word document: {FilePath}", filePath);
                return false;
            }
        }

        public long GetAvailableDiskSpace(string path)
        {
            try
            {
                var pathRoot = Path.GetPathRoot(path);
                if (string.IsNullOrEmpty(pathRoot))
                    throw new ArgumentException($"Cannot determine root path for: {path}");

                var drive = new DriveInfo(pathRoot);
                return drive.AvailableFreeSpace;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting available disk space for path: {Path}", path);
                throw;
            }
        }
    }
}