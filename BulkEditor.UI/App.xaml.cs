using BulkEditor.Application.Services;
using BulkEditor.Core.Configuration;
using BulkEditor.Infrastructure.DependencyInjection;
using BulkEditor.UI.ViewModels;
using Microsoft.Extensions.DependencyInjection;
using Serilog;
using System;
using System.IO;
using System.Windows;

namespace BulkEditor.UI
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        private ServiceProvider? _serviceProvider;

        protected override void OnStartup(StartupEventArgs e)
        {
            try
            {
                // Configure services
                var services = new ServiceCollection();

                // Create default configuration
                var appSettings = CreateDefaultSettings();
                services.AddSingleton(appSettings);

                // Register infrastructure services (simplified)
                services.AddSingleton(Log.Logger);
                services.AddInfrastructureServicesSimplified();

                // Register application services
                services.AddScoped<IApplicationService, ApplicationService>();

                // Register UI Services
                services.AddSingleton<BulkEditor.UI.Services.INotificationService, BulkEditor.UI.Services.NotificationService>();
                services.AddSingleton<BulkEditor.Core.Interfaces.IThemeService, BulkEditor.UI.Services.ThemeService>();

                // Register ViewModels and Views
                services.AddTransient<MainWindowViewModel>();
                services.AddTransient<SettingsViewModel>();
                services.AddTransient<MainWindow>();
                services.AddTransient<BulkEditor.UI.Views.SettingsWindow>();

                // Build service provider
                _serviceProvider = services.BuildServiceProvider();

                // Initialize Serilog
                ConfigureSerilog();

                // Create and show the main window
                var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
                mainWindow.Show();

                base.OnStartup(e);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Application startup failed: {ex.Message}", "Startup Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);

                Log.Fatal(ex, "Application startup failed");
                Shutdown(1);
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            try
            {
                _serviceProvider?.Dispose();
                Log.CloseAndFlush();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error during application shutdown");
            }
            finally
            {
                base.OnExit(e);
            }
        }

        private void Application_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            Log.Fatal(e.Exception, "Unhandled application exception");

            MessageBox.Show($"An unexpected error occurred: {e.Exception.Message}", "Unexpected Error",
                MessageBoxButton.OK, MessageBoxImage.Error);

            e.Handled = true;
            Shutdown(1);
        }

        private AppSettings CreateDefaultSettings()
        {
            return new AppSettings
            {
                Processing = new ProcessingSettings
                {
                    MaxConcurrentDocuments = 200,
                    BatchSize = 50,
                    TimeoutPerDocument = TimeSpan.FromMinutes(5),
                    CreateBackupBeforeProcessing = true,
                    ValidateHyperlinks = true,
                    UpdateHyperlinks = true,
                    AddContentIds = true,
                    OptimizeText = false,
                    LookupIdPattern = @"(TSRC-[^-]+-[0-9]{6}|CMS-[^-]+-[0-9]{6})"
                },
                Validation = new ValidationSettings
                {
                    HttpTimeout = TimeSpan.FromSeconds(30),
                    MaxRetryAttempts = 3,
                    RetryDelay = TimeSpan.FromSeconds(2),
                    UserAgent = "BulkEditor/1.0",
                    CheckExpiredContent = true,
                    FollowRedirects = true,
                    AutoReplaceTitles = false,
                    ReportTitleDifferences = true
                },
                Backup = new BackupSettings
                {
                    BackupDirectory = "Backups",
                    CreateTimestampedBackups = true,
                    MaxBackupAge = 30,
                    CompressBackups = false,
                    AutoCleanupOldBackups = true
                },
                Logging = new LoggingSettings
                {
                    LogLevel = "Information",
                    LogDirectory = "Logs",
                    EnableFileLogging = true,
                    EnableConsoleLogging = true,
                    MaxLogFileSizeMB = 10,
                    MaxLogFiles = 5,
                    LogFilePattern = "bulkeditor-{Date}.log"
                },
                Replacement = new ReplacementSettings
                {
                    EnableHyperlinkReplacement = false,
                    EnableTextReplacement = false,
                    MaxReplacementRules = 50,
                    ValidateContentIds = true
                }
            };
        }

        private void ConfigureSerilog()
        {
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .Enrich.FromLogContext()
                .WriteTo.Console(outputTemplate: "[{Timestamp:HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}")
                .WriteTo.File(
                    path: Path.Combine("Logs", "bulkeditor-.log"),
                    rollingInterval: RollingInterval.Day,
                    outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} {Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            Log.Information("Application starting up");
        }
    }

    /// <summary>
    /// Simplified extension methods for DI registration
    /// </summary>
    public static class ServiceExtensions
    {
        public static IServiceCollection AddInfrastructureServicesSimplified(this IServiceCollection services)
        {
            // Core Services
            services.AddSingleton<BulkEditor.Core.Interfaces.ILoggingService, BulkEditor.Infrastructure.Services.SerilogService>();
            services.AddSingleton<BulkEditor.Core.Interfaces.IFileService, BulkEditor.Infrastructure.Services.FileService>();
            services.AddSingleton<System.Net.Http.HttpClient>();
            services.AddSingleton<BulkEditor.Core.Interfaces.IHttpService, BulkEditor.Infrastructure.Services.HttpService>();
            services.AddSingleton<BulkEditor.Core.Interfaces.ICacheService, BulkEditor.Infrastructure.Services.MemoryCacheService>();

            // Document Processing Services
            services.AddScoped<BulkEditor.Core.Interfaces.IDocumentProcessor, BulkEditor.Infrastructure.Services.DocumentProcessor>();
            services.AddScoped<BulkEditor.Core.Interfaces.IHyperlinkValidator, BulkEditor.Infrastructure.Services.HyperlinkValidator>();
            services.AddScoped<BulkEditor.Core.Interfaces.ITextOptimizer, BulkEditor.Infrastructure.Services.TextOptimizer>();

            // Replacement Services
            services.AddScoped<BulkEditor.Core.Interfaces.IReplacementService, BulkEditor.Infrastructure.Services.ReplacementService>();
            services.AddScoped<BulkEditor.Core.Interfaces.IHyperlinkReplacementService, BulkEditor.Infrastructure.Services.HyperlinkReplacementService>();
            services.AddScoped<BulkEditor.Core.Interfaces.ITextReplacementService, BulkEditor.Infrastructure.Services.TextReplacementService>();

            return services;
        }
    }
}
