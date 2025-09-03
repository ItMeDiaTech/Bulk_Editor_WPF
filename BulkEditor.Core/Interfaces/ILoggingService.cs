using System;

namespace BulkEditor.Core.Interfaces
{
    /// <summary>
    /// Interface for logging operations
    /// </summary>
    public interface ILoggingService
    {
        /// <summary>
        /// Logs information message
        /// </summary>
        void LogInformation(string message);

        /// <summary>
        /// Logs information message with parameters
        /// </summary>
        void LogInformation(string template, params object[] args);

        /// <summary>
        /// Logs warning message
        /// </summary>
        void LogWarning(string message);

        /// <summary>
        /// Logs warning message with parameters
        /// </summary>
        void LogWarning(string template, params object[] args);

        /// <summary>
        /// Logs error message
        /// </summary>
        void LogError(string message);

        /// <summary>
        /// Logs error message with exception
        /// </summary>
        void LogError(Exception exception, string message);

        /// <summary>
        /// Logs error message with parameters
        /// </summary>
        void LogError(string template, params object[] args);

        /// <summary>
        /// Logs error with exception and parameters
        /// </summary>
        void LogError(Exception exception, string template, params object[] args);

        /// <summary>
        /// Logs debug message
        /// </summary>
        void LogDebug(string message);

        /// <summary>
        /// Logs debug message with parameters
        /// </summary>
        void LogDebug(string template, params object[] args);

        /// <summary>
        /// Logs verbose/trace message
        /// </summary>
        void LogVerbose(string message);

        /// <summary>
        /// Logs verbose message with parameters
        /// </summary>
        void LogVerbose(string template, params object[] args);
    }
}