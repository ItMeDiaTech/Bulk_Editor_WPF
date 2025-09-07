using BulkEditor.Core.Configuration;
using BulkEditor.Core.Interfaces;
using BulkEditor.Infrastructure.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Serilog;
using System;
using System.IO;
using System.Net.Http;

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

            // CRITICAL FIX: HTTP Services with proper configuration (Issues #22-23)
            services.AddSingleton<HttpClient>(provider =>
            {
                var handler = new HttpClientHandler()
                {
                    // CRITICAL FIX: Connection pooling configuration (Issue #22)
                    MaxConnectionsPerServer = 10,
                    UseProxy = false, // Disable proxy for corporate networks if needed
                    UseCookies = false // Disable cookies for stateless API calls
                };

                var client = new HttpClient(handler);

                // CRITICAL FIX: Proper timeout configuration (Issue #23)
                client.Timeout = TimeSpan.FromMinutes(5); // 5 minute timeout for API calls

                // CRITICAL FIX: Proper User-Agent for API compatibility
                client.DefaultRequestHeaders.Add("User-Agent", "BulkEditor-WPF/1.0");

                // CRITICAL FIX: Accept headers for API compatibility
                client.DefaultRequestHeaders.Add("Accept", "application/json");
                client.DefaultRequestHeaders.Add("Accept-Charset", "utf-8");

                return client;
            });
            services.AddSingleton<IHttpService, HttpService>();

            // Document Processing Services
            services.AddScoped<IDocumentProcessor, DocumentProcessor>();
            services.AddScoped<IHyperlinkValidator, HyperlinkValidator>();
            services.AddScoped<IHyperlinkReplacementService, HyperlinkReplacementService>();
            services.AddScoped<IReplacementService, ReplacementService>();
            services.AddScoped<ITextOptimizer, TextOptimizer>();
            services.AddScoped<ICacheService, MemoryCacheService>();
            services.AddScoped<BulkEditor.Core.Services.IConfigurationService, ConfigurationService>();
            services.AddScoped<BulkEditor.Core.Services.IUpdateService, GitHubUpdateService>();

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
