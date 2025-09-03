using BulkEditor.Core.Interfaces;
using Serilog;
using System;

namespace BulkEditor.Infrastructure.Services
{
    /// <summary>
    /// Serilog implementation of the logging service
    /// </summary>
    public class SerilogService : ILoggingService
    {
        private readonly ILogger _logger;

        public SerilogService()
        {
            _logger = Log.Logger;
        }

        public void LogInformation(string message)
        {
            _logger.Information(message);
        }

        public void LogInformation(string template, params object[] args)
        {
            _logger.Information(template, args);
        }

        public void LogWarning(string message)
        {
            _logger.Warning(message);
        }

        public void LogWarning(string template, params object[] args)
        {
            _logger.Warning(template, args);
        }

        public void LogError(string message)
        {
            _logger.Error(message);
        }

        public void LogError(Exception exception, string message)
        {
            _logger.Error(exception, message);
        }

        public void LogError(string template, params object[] args)
        {
            _logger.Error(template, args);
        }

        public void LogError(Exception exception, string template, params object[] args)
        {
            _logger.Error(exception, template, args);
        }

        public void LogDebug(string message)
        {
            _logger.Debug(message);
        }

        public void LogDebug(string template, params object[] args)
        {
            _logger.Debug(template, args);
        }

        public void LogVerbose(string message)
        {
            _logger.Verbose(message);
        }

        public void LogVerbose(string template, params object[] args)
        {
            _logger.Verbose(template, args);
        }
    }
}