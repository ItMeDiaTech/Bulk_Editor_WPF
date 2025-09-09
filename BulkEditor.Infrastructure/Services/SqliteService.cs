using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using BulkEditor.Core.Models;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// SQLite implementation of database service for persistent storage
    /// </summary>
    public class SqliteService : IDatabaseService, IDisposable
    {
        private readonly ILoggingService _logger;
        private readonly string _connectionString;
        private readonly string _databasePath;
        private bool _initialized;
        private readonly SemaphoreSlim _initLock = new(1, 1);

        public SqliteService(ILoggingService logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            
            // Store database in AppData/BulkEditor/Database/
            var appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "BulkEditor", "Database");
            
            Directory.CreateDirectory(appDataPath);
            
            _databasePath = Path.Combine(appDataPath, "BulkEditor.db");
            _connectionString = $"Data Source={_databasePath};Foreign Keys=True;";
        }

        public async Task InitializeAsync(CancellationToken cancellationToken = default)
        {
            if (_initialized) return;

            await _initLock.WaitAsync(cancellationToken);
            try
            {
                if (_initialized) return;

                _logger.LogInformation("Initializing SQLite database at: {DatabasePath}", _databasePath);

                using var connection = new SqliteConnection(_connectionString);
                await connection.OpenAsync(cancellationToken);

                await CreateTablesAsync(connection, cancellationToken);
                
                _initialized = true;
                _logger.LogInformation("SQLite database initialized successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to initialize SQLite database");
                throw;
            }
            finally
            {
                _initLock.Release();
            }
        }

        private async Task CreateTablesAsync(SqliteConnection connection, CancellationToken cancellationToken)
        {
            // Settings table
            await ExecuteNonQueryAsync(connection, @"
                CREATE TABLE IF NOT EXISTS Settings (
                    Key TEXT NOT NULL,
                    Category TEXT NOT NULL DEFAULT 'General',
                    Value TEXT NOT NULL,
                    CreatedAt DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    UpdatedAt DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY (Key, Category)
                )", cancellationToken);

            // Processing Sessions table
            await ExecuteNonQueryAsync(connection, @"
                CREATE TABLE IF NOT EXISTS ProcessingSessions (
                    SessionId TEXT PRIMARY KEY,
                    StartTime DATETIME NOT NULL,
                    EndTime DATETIME,
                    TotalDocuments INTEGER NOT NULL DEFAULT 0,
                    ProcessedDocuments INTEGER NOT NULL DEFAULT 0,
                    SuccessfulDocuments INTEGER NOT NULL DEFAULT 0,
                    FailedDocuments INTEGER NOT NULL DEFAULT 0,
                    TotalProcessingTimeMs INTEGER,
                    Status TEXT NOT NULL DEFAULT 'Running',
                    ErrorMessage TEXT,
                    Metadata TEXT
                )", cancellationToken);

            // Document Processing Results table
            await ExecuteNonQueryAsync(connection, @"
                CREATE TABLE IF NOT EXISTS DocumentProcessingResults (
                    Id TEXT PRIMARY KEY,
                    SessionId TEXT NOT NULL,
                    DocumentPath TEXT NOT NULL,
                    DocumentName TEXT NOT NULL,
                    FileSizeBytes INTEGER NOT NULL,
                    ProcessingStartTime DATETIME NOT NULL,
                    ProcessingEndTime DATETIME NOT NULL,
                    ProcessingDurationMs INTEGER NOT NULL,
                    IsSuccessful BOOLEAN NOT NULL,
                    ErrorMessage TEXT,
                    HyperlinksProcessed INTEGER NOT NULL DEFAULT 0,
                    HyperlinksUpdated INTEGER NOT NULL DEFAULT 0,
                    TextReplacements INTEGER NOT NULL DEFAULT 0,
                    Metadata TEXT,
                    FOREIGN KEY (SessionId) REFERENCES ProcessingSessions(SessionId) ON DELETE CASCADE
                )", cancellationToken);

            // Performance Metrics table
            await ExecuteNonQueryAsync(connection, @"
                CREATE TABLE IF NOT EXISTS PerformanceMetrics (
                    Id TEXT PRIMARY KEY,
                    OperationName TEXT NOT NULL,
                    Timestamp DATETIME NOT NULL,
                    DurationMs INTEGER NOT NULL,
                    MemoryUsedBytes INTEGER NOT NULL,
                    ThreadId INTEGER NOT NULL,
                    MachineName TEXT NOT NULL,
                    CustomMetrics TEXT
                )", cancellationToken);

            // Cache Entries table
            await ExecuteNonQueryAsync(connection, @"
                CREATE TABLE IF NOT EXISTS CacheEntries (
                    Key TEXT PRIMARY KEY,
                    Value TEXT NOT NULL,
                    CreatedAt DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    ExpiryDate DATETIME,
                    AccessCount INTEGER NOT NULL DEFAULT 0,
                    LastAccessedAt DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
                )", cancellationToken);

            // Create indexes for better query performance
            await ExecuteNonQueryAsync(connection, "CREATE INDEX IF NOT EXISTS IX_Settings_Category ON Settings(Category)", cancellationToken);
            await ExecuteNonQueryAsync(connection, "CREATE INDEX IF NOT EXISTS IX_ProcessingSessions_StartTime ON ProcessingSessions(StartTime)", cancellationToken);
            await ExecuteNonQueryAsync(connection, "CREATE INDEX IF NOT EXISTS IX_DocumentResults_SessionId ON DocumentProcessingResults(SessionId)", cancellationToken);
            await ExecuteNonQueryAsync(connection, "CREATE INDEX IF NOT EXISTS IX_DocumentResults_DocumentPath ON DocumentProcessingResults(DocumentPath)", cancellationToken);
            await ExecuteNonQueryAsync(connection, "CREATE INDEX IF NOT EXISTS IX_PerformanceMetrics_OperationName ON PerformanceMetrics(OperationName)", cancellationToken);
            await ExecuteNonQueryAsync(connection, "CREATE INDEX IF NOT EXISTS IX_PerformanceMetrics_Timestamp ON PerformanceMetrics(Timestamp)", cancellationToken);
            await ExecuteNonQueryAsync(connection, "CREATE INDEX IF NOT EXISTS IX_CacheEntries_ExpiryDate ON CacheEntries(ExpiryDate)", cancellationToken);
        }

        public async Task SaveSettingsAsync(string key, string value, string category = "General", CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            await ExecuteNonQueryAsync(connection, @"
                INSERT INTO Settings (Key, Category, Value, UpdatedAt) 
                VALUES (@key, @category, @value, CURRENT_TIMESTAMP)
                ON CONFLICT(Key, Category) DO UPDATE SET 
                    Value = @value, 
                    UpdatedAt = CURRENT_TIMESTAMP",
                cancellationToken,
                ("@key", key),
                ("@category", category),
                ("@value", value));
        }

        public async Task<string?> GetSettingsAsync(string key, string category = "General", CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            using var command = connection.CreateCommand();
            command.CommandText = "SELECT Value FROM Settings WHERE Key = @key AND Category = @category";
            command.Parameters.AddWithValue("@key", key);
            command.Parameters.AddWithValue("@category", category);

            var result = await command.ExecuteScalarAsync(cancellationToken);
            return result?.ToString();
        }

        public async Task<Dictionary<string, string>> GetAllSettingsAsync(string category = "General", CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            var settings = new Dictionary<string, string>();

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            using var command = connection.CreateCommand();
            command.CommandText = "SELECT Key, Value FROM Settings WHERE Category = @category ORDER BY Key";
            command.Parameters.AddWithValue("@category", category);

            using var reader = await command.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                settings[reader.GetString(0)] = reader.GetString(1);
            }

            return settings;
        }

        public async Task SaveProcessingSessionAsync(ProcessingSession session, CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            var metadataJson = JsonSerializer.Serialize(session.Metadata);
            var totalProcessingTimeMs = session.TotalProcessingTime?.TotalMilliseconds;

            await ExecuteNonQueryAsync(connection, @"
                INSERT OR REPLACE INTO ProcessingSessions 
                (SessionId, StartTime, EndTime, TotalDocuments, ProcessedDocuments, 
                 SuccessfulDocuments, FailedDocuments, TotalProcessingTimeMs, Status, ErrorMessage, Metadata)
                VALUES (@sessionId, @startTime, @endTime, @totalDocuments, @processedDocuments,
                        @successfulDocuments, @failedDocuments, @totalProcessingTimeMs, @status, @errorMessage, @metadata)",
                cancellationToken,
                ("@sessionId", session.SessionId.ToString()),
                ("@startTime", session.StartTime),
                ("@endTime", session.EndTime),
                ("@totalDocuments", session.TotalDocuments),
                ("@processedDocuments", session.ProcessedDocuments),
                ("@successfulDocuments", session.SuccessfulDocuments),
                ("@failedDocuments", session.FailedDocuments),
                ("@totalProcessingTimeMs", totalProcessingTimeMs),
                ("@status", session.Status),
                ("@errorMessage", session.ErrorMessage),
                ("@metadata", metadataJson));
        }

        public async Task<ProcessingSession?> GetProcessingSessionAsync(Guid sessionId, CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            using var command = connection.CreateCommand();
            command.CommandText = @"
                SELECT SessionId, StartTime, EndTime, TotalDocuments, ProcessedDocuments, 
                       SuccessfulDocuments, FailedDocuments, TotalProcessingTimeMs, Status, ErrorMessage, Metadata
                FROM ProcessingSessions WHERE SessionId = @sessionId";
            command.Parameters.AddWithValue("@sessionId", sessionId.ToString());

            using var reader = await command.ExecuteReaderAsync(cancellationToken);
            if (await reader.ReadAsync(cancellationToken))
            {
                return CreateProcessingSessionFromReader(reader);
            }

            return null;
        }

        public async Task<IEnumerable<ProcessingSession>> GetRecentProcessingSessionsAsync(int limit = 50, CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            var sessions = new List<ProcessingSession>();

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            using var command = connection.CreateCommand();
            command.CommandText = @"
                SELECT SessionId, StartTime, EndTime, TotalDocuments, ProcessedDocuments, 
                       SuccessfulDocuments, FailedDocuments, TotalProcessingTimeMs, Status, ErrorMessage, Metadata
                FROM ProcessingSessions 
                ORDER BY StartTime DESC 
                LIMIT @limit";
            command.Parameters.AddWithValue("@limit", limit);

            using var reader = await command.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                sessions.Add(CreateProcessingSessionFromReader(reader));
            }

            return sessions;
        }

        public async Task SaveDocumentProcessingResultAsync(DocumentProcessingResult result, CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            var metadataJson = JsonSerializer.Serialize(result.Metadata);

            await ExecuteNonQueryAsync(connection, @"
                INSERT OR REPLACE INTO DocumentProcessingResults
                (Id, SessionId, DocumentPath, DocumentName, FileSizeBytes, ProcessingStartTime,
                 ProcessingEndTime, ProcessingDurationMs, IsSuccessful, ErrorMessage, 
                 HyperlinksProcessed, HyperlinksUpdated, TextReplacements, Metadata)
                VALUES (@id, @sessionId, @documentPath, @documentName, @fileSizeBytes, @processingStartTime,
                        @processingEndTime, @processingDurationMs, @isSuccessful, @errorMessage,
                        @hyperlinksProcessed, @hyperlinksUpdated, @textReplacements, @metadata)",
                cancellationToken,
                ("@id", result.Id.ToString()),
                ("@sessionId", result.SessionId.ToString()),
                ("@documentPath", result.DocumentPath),
                ("@documentName", result.DocumentName),
                ("@fileSizeBytes", result.FileSizeBytes),
                ("@processingStartTime", result.ProcessingStartTime),
                ("@processingEndTime", result.ProcessingEndTime),
                ("@processingDurationMs", (int)result.ProcessingDuration.TotalMilliseconds),
                ("@isSuccessful", result.IsSuccessful),
                ("@errorMessage", result.ErrorMessage),
                ("@hyperlinksProcessed", result.HyperlinksProcessed),
                ("@hyperlinksUpdated", result.HyperlinksUpdated),
                ("@textReplacements", result.TextReplacements),
                ("@metadata", metadataJson));
        }

        public async Task<IEnumerable<DocumentProcessingResult>> GetDocumentProcessingHistoryAsync(string? documentPath = null, int limit = 100, CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            var results = new List<DocumentProcessingResult>();

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            using var command = connection.CreateCommand();
            if (string.IsNullOrEmpty(documentPath))
            {
                command.CommandText = @"
                    SELECT Id, SessionId, DocumentPath, DocumentName, FileSizeBytes, ProcessingStartTime,
                           ProcessingEndTime, ProcessingDurationMs, IsSuccessful, ErrorMessage,
                           HyperlinksProcessed, HyperlinksUpdated, TextReplacements, Metadata
                    FROM DocumentProcessingResults 
                    ORDER BY ProcessingStartTime DESC 
                    LIMIT @limit";
            }
            else
            {
                command.CommandText = @"
                    SELECT Id, SessionId, DocumentPath, DocumentName, FileSizeBytes, ProcessingStartTime,
                           ProcessingEndTime, ProcessingDurationMs, IsSuccessful, ErrorMessage,
                           HyperlinksProcessed, HyperlinksUpdated, TextReplacements, Metadata
                    FROM DocumentProcessingResults 
                    WHERE DocumentPath = @documentPath
                    ORDER BY ProcessingStartTime DESC 
                    LIMIT @limit";
                command.Parameters.AddWithValue("@documentPath", documentPath);
            }
            command.Parameters.AddWithValue("@limit", limit);

            using var reader = await command.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                results.Add(CreateDocumentProcessingResultFromReader(reader));
            }

            return results;
        }

        public async Task SavePerformanceMetricAsync(PerformanceMetric metric, CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            var customMetricsJson = JsonSerializer.Serialize(metric.CustomMetrics);

            await ExecuteNonQueryAsync(connection, @"
                INSERT INTO PerformanceMetrics
                (Id, OperationName, Timestamp, DurationMs, MemoryUsedBytes, ThreadId, MachineName, CustomMetrics)
                VALUES (@id, @operationName, @timestamp, @durationMs, @memoryUsedBytes, @threadId, @machineName, @customMetrics)",
                cancellationToken,
                ("@id", metric.Id.ToString()),
                ("@operationName", metric.OperationName),
                ("@timestamp", metric.Timestamp),
                ("@durationMs", (int)metric.Duration.TotalMilliseconds),
                ("@memoryUsedBytes", metric.MemoryUsedBytes),
                ("@threadId", metric.ThreadId),
                ("@machineName", metric.MachineName),
                ("@customMetrics", customMetricsJson));
        }

        public async Task<IEnumerable<PerformanceMetric>> GetPerformanceMetricsAsync(string? operationName = null, DateTime? fromDate = null, int limit = 1000, CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            var metrics = new List<PerformanceMetric>();

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            using var command = connection.CreateCommand();
            var whereClause = new List<string>();
            
            if (!string.IsNullOrEmpty(operationName))
            {
                whereClause.Add("OperationName = @operationName");
                command.Parameters.AddWithValue("@operationName", operationName);
            }
            
            if (fromDate.HasValue)
            {
                whereClause.Add("Timestamp >= @fromDate");
                command.Parameters.AddWithValue("@fromDate", fromDate.Value);
            }

            var whereClauseSql = whereClause.Count > 0 ? "WHERE " + string.Join(" AND ", whereClause) : "";

            command.CommandText = $@"
                SELECT Id, OperationName, Timestamp, DurationMs, MemoryUsedBytes, ThreadId, MachineName, CustomMetrics
                FROM PerformanceMetrics
                {whereClauseSql}
                ORDER BY Timestamp DESC
                LIMIT @limit";
            command.Parameters.AddWithValue("@limit", limit);

            using var reader = await command.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                metrics.Add(CreatePerformanceMetricFromReader(reader));
            }

            return metrics;
        }

        public async Task SaveCacheEntryAsync(string key, string value, DateTime? expiryDate = null, CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            await ExecuteNonQueryAsync(connection, @"
                INSERT OR REPLACE INTO CacheEntries (Key, Value, ExpiryDate, AccessCount, LastAccessedAt)
                VALUES (@key, @value, @expiryDate, 1, CURRENT_TIMESTAMP)",
                cancellationToken,
                ("@key", key),
                ("@value", value),
                ("@expiryDate", expiryDate));
        }

        public async Task<string?> GetCacheEntryAsync(string key, CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            // Check if entry exists and is not expired
            using var checkCommand = connection.CreateCommand();
            checkCommand.CommandText = @"
                SELECT Value FROM CacheEntries 
                WHERE Key = @key AND (ExpiryDate IS NULL OR ExpiryDate > CURRENT_TIMESTAMP)";
            checkCommand.Parameters.AddWithValue("@key", key);

            var result = await checkCommand.ExecuteScalarAsync(cancellationToken);
            
            if (result != null)
            {
                // Update access count and last accessed time
                using var updateCommand = connection.CreateCommand();
                updateCommand.CommandText = @"
                    UPDATE CacheEntries 
                    SET AccessCount = AccessCount + 1, LastAccessedAt = CURRENT_TIMESTAMP 
                    WHERE Key = @key";
                updateCommand.Parameters.AddWithValue("@key", key);
                await updateCommand.ExecuteNonQueryAsync(cancellationToken);
            }

            return result?.ToString();
        }

        public async Task CleanupExpiredCacheAsync(CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            var deletedCount = await ExecuteNonQueryAsync(connection, 
                "DELETE FROM CacheEntries WHERE ExpiryDate IS NOT NULL AND ExpiryDate <= CURRENT_TIMESTAMP", 
                cancellationToken);

            if (deletedCount > 0)
            {
                _logger.LogInformation("Cleaned up {Count} expired cache entries", deletedCount);
            }
        }

        public async Task<DatabaseStatistics> GetStatisticsAsync(CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            var stats = new DatabaseStatistics();

            // Get database size
            var dbFileInfo = new FileInfo(_databasePath);
            stats.DatabaseSizeBytes = dbFileInfo.Exists ? dbFileInfo.Length : 0;

            // Get table counts
            stats.SettingsCount = await GetTableCountAsync(connection, "Settings", cancellationToken);
            stats.ProcessingSessionsCount = await GetTableCountAsync(connection, "ProcessingSessions", cancellationToken);
            stats.DocumentResultsCount = await GetTableCountAsync(connection, "DocumentProcessingResults", cancellationToken);
            stats.PerformanceMetricsCount = await GetTableCountAsync(connection, "PerformanceMetrics", cancellationToken);
            stats.CacheEntriesCount = await GetTableCountAsync(connection, "CacheEntries", cancellationToken);

            stats.LastMaintenanceDate = DateTime.UtcNow; // Placeholder - could be stored in settings

            return stats;
        }

        public async Task PerformMaintenanceAsync(CancellationToken cancellationToken = default)
        {
            await EnsureInitializedAsync(cancellationToken);

            _logger.LogInformation("Starting database maintenance");

            using var connection = new SqliteConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            // Clean expired cache entries
            await CleanupExpiredCacheAsync(cancellationToken);

            // Clean old performance metrics (keep last 30 days)
            var thirtyDaysAgo = DateTime.UtcNow.AddDays(-30);
            var deletedMetrics = await ExecuteNonQueryAsync(connection,
                "DELETE FROM PerformanceMetrics WHERE Timestamp < @cutoffDate",
                cancellationToken,
                ("@cutoffDate", thirtyDaysAgo));

            // Clean old processing sessions (keep last 90 days)
            var ninetyDaysAgo = DateTime.UtcNow.AddDays(-90);
            var deletedSessions = await ExecuteNonQueryAsync(connection,
                "DELETE FROM ProcessingSessions WHERE StartTime < @cutoffDate",
                cancellationToken,
                ("@cutoffDate", ninetyDaysAgo));

            // Vacuum database to reclaim space
            await ExecuteNonQueryAsync(connection, "VACUUM", cancellationToken);

            // Analyze tables for better query performance
            await ExecuteNonQueryAsync(connection, "ANALYZE", cancellationToken);

            _logger.LogInformation("Database maintenance completed. Deleted {MetricsCount} old metrics and {SessionsCount} old sessions",
                deletedMetrics, deletedSessions);
        }

        private ProcessingSession CreateProcessingSessionFromReader(SqliteDataReader reader)
        {
            var session = new ProcessingSession
            {
                SessionId = Guid.Parse(reader.GetString(0)), // SessionId
                StartTime = reader.GetDateTime(1), // StartTime
                EndTime = reader.IsDBNull(2) ? null : reader.GetDateTime(2), // EndTime
                TotalDocuments = reader.GetInt32(3), // TotalDocuments
                ProcessedDocuments = reader.GetInt32(4), // ProcessedDocuments
                SuccessfulDocuments = reader.GetInt32(5), // SuccessfulDocuments
                FailedDocuments = reader.GetInt32(6), // FailedDocuments
                Status = reader.GetString(9), // Status
                ErrorMessage = reader.IsDBNull(10) ? null : reader.GetString(10) // ErrorMessage
            };

            if (!reader.IsDBNull(7)) // TotalProcessingTimeMs
            {
                session.TotalProcessingTime = TimeSpan.FromMilliseconds(reader.GetDouble(7));
            }

            var metadataJson = reader.IsDBNull(11) ? null : reader.GetString(11); // Metadata
            if (!string.IsNullOrEmpty(metadataJson))
            {
                try
                {
                    session.Metadata = JsonSerializer.Deserialize<Dictionary<string, object>>(metadataJson) ?? new();
                }
                catch (Exception ex)
                {
                    _logger.LogWarning("Failed to deserialize session metadata: {ErrorMessage}", ex.Message);
                }
            }

            return session;
        }

        private DocumentProcessingResult CreateDocumentProcessingResultFromReader(SqliteDataReader reader)
        {
            var result = new DocumentProcessingResult
            {
                Id = Guid.Parse(reader.GetString(0)), // Id
                SessionId = Guid.Parse(reader.GetString(1)), // SessionId
                DocumentPath = reader.GetString(2), // DocumentPath
                DocumentName = reader.GetString(3), // DocumentName
                FileSizeBytes = reader.GetInt64(4), // FileSizeBytes
                ProcessingStartTime = reader.GetDateTime(5), // ProcessingStartTime
                ProcessingEndTime = reader.GetDateTime(6), // ProcessingEndTime
                ProcessingDuration = TimeSpan.FromMilliseconds(reader.GetInt32(7)), // ProcessingDurationMs
                IsSuccessful = reader.GetBoolean(8), // IsSuccessful
                ErrorMessage = reader.IsDBNull(9) ? null : reader.GetString(9), // ErrorMessage
                HyperlinksProcessed = reader.GetInt32(10), // HyperlinksProcessed
                HyperlinksUpdated = reader.GetInt32(11), // HyperlinksUpdated
                TextReplacements = reader.GetInt32(12) // TextReplacements
            };

            var metadataJson = reader.IsDBNull(13) ? null : reader.GetString(13); // Metadata
            if (!string.IsNullOrEmpty(metadataJson))
            {
                try
                {
                    result.Metadata = JsonSerializer.Deserialize<Dictionary<string, object>>(metadataJson) ?? new();
                }
                catch (Exception ex)
                {
                    _logger.LogWarning("Failed to deserialize document result metadata: {ErrorMessage}", ex.Message);
                }
            }

            return result;
        }

        private PerformanceMetric CreatePerformanceMetricFromReader(SqliteDataReader reader)
        {
            var metric = new PerformanceMetric
            {
                Id = Guid.Parse(reader.GetString(0)), // Id
                OperationName = reader.GetString(1), // OperationName
                Timestamp = reader.GetDateTime(2), // Timestamp
                Duration = TimeSpan.FromMilliseconds(reader.GetInt32(3)), // DurationMs
                MemoryUsedBytes = reader.GetInt64(4), // MemoryUsedBytes
                ThreadId = reader.GetInt32(5), // ThreadId
                MachineName = reader.GetString(6) // MachineName
            };

            var customMetricsJson = reader.IsDBNull(7) ? null : reader.GetString(7); // CustomMetrics
            if (!string.IsNullOrEmpty(customMetricsJson))
            {
                try
                {
                    metric.CustomMetrics = JsonSerializer.Deserialize<Dictionary<string, object>>(customMetricsJson) ?? new();
                }
                catch (Exception ex)
                {
                    _logger.LogWarning("Failed to deserialize performance metric custom metrics: {ErrorMessage}", ex.Message);
                }
            }

            return metric;
        }

        private async Task<int> GetTableCountAsync(SqliteConnection connection, string tableName, CancellationToken cancellationToken)
        {
            using var command = connection.CreateCommand();
            command.CommandText = $"SELECT COUNT(*) FROM {tableName}";
            var result = await command.ExecuteScalarAsync(cancellationToken);
            return Convert.ToInt32(result);
        }

        private async Task<int> ExecuteNonQueryAsync(SqliteConnection connection, string sql, CancellationToken cancellationToken, params (string name, object? value)[] parameters)
        {
            using var command = connection.CreateCommand();
            command.CommandText = sql;
            
            foreach (var (name, value) in parameters)
            {
                command.Parameters.AddWithValue(name, value ?? DBNull.Value);
            }

            return await command.ExecuteNonQueryAsync(cancellationToken);
        }

        private async Task EnsureInitializedAsync(CancellationToken cancellationToken)
        {
            if (!_initialized)
            {
                await InitializeAsync(cancellationToken);
            }
        }

        public void Dispose()
        {
            _initLock?.Dispose();
        }
    }
}