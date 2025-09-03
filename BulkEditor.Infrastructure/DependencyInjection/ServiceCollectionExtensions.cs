using BulkEditor.Core.Configuration;
using BulkEditor.Core.Interfaces;
using BulkEditor.Infrastructure.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Serilog;
using System;
using System.IO;

namespace BulkEditor.Infrastructure.DependencyInjection
{
    /// <summary>
    /// Extension methods for configuring services in the DI container
    /// </summary>
    public static class ServiceCollectionExtensions
    {
        /// <summary>
        /// Adds all infrastructure services to the service collection
        /// </summary>
        public static IServiceCollection AddInfrastructureServices(this IServiceCollection services, IConfiguration configuration)
        {
            // Configuration
            services.Configure<AppSettings>(configuration);

            // Core Services
            services.AddSingleton<ILoggingService, SerilogService>();
            services.AddSingleton<IFileService, FileService>();
            services.AddSingleton<HttpClient>();
            services.AddSingleton<IHttpService, HttpService>();

            // Document Processing Services
            services.AddScoped<IDocumentProcessor, DocumentProcessor>();
            services.AddScoped<IHyperlinkValidator, HyperlinkValidator>();

            // Configure Serilog
            ConfigureSerilog(configuration);

            return services;
        }

        /// <summary>
        /// Configures Serilog logging
        /// </summary>
        private static void ConfigureSerilog(IConfiguration configuration)
        {
            var loggingSettings = configuration.GetSection("Logging").Get<LoggingSettings>() ?? new LoggingSettings();

            var loggerConfig = new LoggerConfiguration()
                .MinimumLevel.Is(GetLogEventLevel(loggingSettings.LogLevel))
                .Enrich.FromLogContext();

            if (loggingSettings.EnableConsoleLogging)
            {
                loggerConfig.WriteTo.Console(
                    outputTemplate: "[{Timestamp:HH:mm:ss} {Level:u3}] {Message:lj} {NewLine}{Exception}");
            }

            if (loggingSettings.EnableFileLogging)
            {
                var logPath = Path.Combine(loggingSettings.LogDirectory, loggingSettings.LogFilePattern);
                loggerConfig.WriteTo.File(
                    path: logPath,
                    rollingInterval: RollingInterval.Day,
                    fileSizeLimitBytes: loggingSettings.MaxLogFileSizeMB * 1024 * 1024,
                    retainedFileCountLimit: loggingSettings.MaxLogFiles,
                    outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} {Level:u3}] {Message:lj} {NewLine}{Exception}");
            }

            Log.Logger = loggerConfig.CreateLogger();
        }

        /// <summary>
        /// Converts string log level to Serilog LogEventLevel
        /// </summary>
        private static Serilog.Events.LogEventLevel GetLogEventLevel(string logLevel)
        {
            return logLevel.ToUpperInvariant() switch
            {
                "VERBOSE" => Serilog.Events.LogEventLevel.Verbose,
                "DEBUG" => Serilog.Events.LogEventLevel.Debug,
                "INFORMATION" => Serilog.Events.LogEventLevel.Information,
                "WARNING" => Serilog.Events.LogEventLevel.Warning,
                "ERROR" => Serilog.Events.LogEventLevel.Error,
                "FATAL" => Serilog.Events.LogEventLevel.Fatal,
                _ => Serilog.Events.LogEventLevel.Information
            };
        }
    }
}