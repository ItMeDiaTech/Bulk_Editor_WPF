using BulkEditor.Application.Services;
using BulkEditor.Core.Configuration;
using BulkEditor.Core.Interfaces;
using BulkEditor.Core.Services;
using BulkEditor.Infrastructure.Services;
using BulkEditor.UI.ViewModels;
using Microsoft.Extensions.DependencyInjection;
using Serilog;
using System;
using System.IO;
using System.Net.Http;
using System.Windows;

namespace BulkEditor.UI
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        private ServiceProvider? _serviceProvider;
        private BulkEditor.Application.Services.UpdateManager? _updateManager;

        protected override async void OnStartup(StartupEventArgs e)
        {
            try
            {
                // Create temporary logger for ConfigurationService (minimal configuration)
                var tempLogger = new BulkEditor.Infrastructure.Services.SerilogService();

                // Initialize configuration service early to avoid double registration
                var configService = new ConfigurationService(tempLogger);
                await configService.InitializeAsync();
                await configService.MigrateSettingsAsync();
                var appSettings = await configService.LoadSettingsAsync();

                // CRITICAL FIX: Configure Serilog only ONCE with proper settings
                ConfigureSerilog(appSettings);

                // Configure services
                var services = new ServiceCollection();

                // Register core instances
                services.AddSingleton<IConfigurationService>(configService);
                services.AddSingleton(appSettings);
                
                // CRITICAL FIX: Register IOptions<AppSettings> for dependency injection
                services.AddSingleton<Microsoft.Extensions.Options.IOptions<AppSettings>>(provider =>
                    Microsoft.Extensions.Options.Options.Create(appSettings));

                // Register infrastructure services
                services.AddSingleton<BulkEditor.Core.Interfaces.ILoggingService, BulkEditor.Infrastructure.Services.SerilogService>();
                services.AddSingleton<BulkEditor.Core.Interfaces.IFileService, BulkEditor.Infrastructure.Services.FileService>();

                // Configure HttpClient properly with all necessary headers and settings
                services.AddSingleton<System.Net.Http.HttpClient>(provider =>
                {
                    var httpClient = new HttpClient();
                    httpClient.Timeout = TimeSpan.FromSeconds(30);
                    httpClient.DefaultRequestHeaders.Add("User-Agent", "BulkEditor/1.0");
                    httpClient.DefaultRequestHeaders.Add("Accept", "application/vnd.github.v3+json");
                    return httpClient;
                });

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

                // Register application services
                services.AddScoped<IApplicationService, BulkEditor.Application.Services.ApplicationService>();

                // Register undo/revert services
                services.AddSingleton<ISessionManager, SessionManager>();
                services.AddSingleton<IBackupService, BackupService>();
                services.AddSingleton<IUndoService, UndoService>();

                // Register update services
                services.AddSingleton<IUpdateService, GitHubUpdateService>(provider =>
                {
                    var httpClient = provider.GetRequiredService<System.Net.Http.HttpClient>();
                    var logger = provider.GetRequiredService<BulkEditor.Core.Interfaces.ILoggingService>();
                    var configServiceProvider = provider.GetRequiredService<IConfigurationService>();
                    return new GitHubUpdateService(httpClient, logger, configServiceProvider,
                        appSettings.Update.GitHubOwner, appSettings.Update.GitHubRepository);
                });
                services.AddSingleton<BulkEditor.Application.Services.UpdateManager>();

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

                // Initialize update manager
                _updateManager = _serviceProvider.GetRequiredService<BulkEditor.Application.Services.UpdateManager>();
                await _updateManager.StartAsync();

                // Subscribe to update events
                var updateService = _serviceProvider.GetRequiredService<IUpdateService>();
                if (updateService is GitHubUpdateService githubUpdateService)
                {
                    githubUpdateService.UpdateRequiresRestart += OnUpdateRequiresRestart;
                }

                // Create and show the main window
                var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
                mainWindow.Show();

                base.OnStartup(e);

                // Log successful startup
                Log.Information("Application started successfully");
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
                if (_updateManager != null)
                {
                    _updateManager.Stop();
                    _updateManager.Dispose();
                }
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

        private void OnUpdateRequiresRestart(object sender, EventArgs e)
        {
            Log.Information("Update requires application restart");
            Dispatcher.BeginInvoke(() =>
            {
                Shutdown();
            });
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
                },
                Api = new ApiSettings
                {
                    BaseUrl = string.Empty,
                    ApiKey = string.Empty,
                    Timeout = TimeSpan.FromSeconds(30),
                    EnableCaching = true,
                    CacheExpiry = TimeSpan.FromHours(1)
                },
                Update = new UpdateSettings
                {
                    AutoUpdateEnabled = true,
                    CheckIntervalHours = 24,
                    InstallSecurityUpdatesAutomatically = true,
                    NotifyOnUpdatesAvailable = true,
                    CreateBackupBeforeUpdate = true,
                    GitHubOwner = "DiaTech",
                    GitHubRepository = "Bulk_Editor",
                    IncludePrerelease = false
                }
            };
        }

        private void ConfigureSerilog(AppSettings appSettings = null)
        {
            var logDirectory = appSettings?.Logging?.LogDirectory ??
                              Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BulkEditor", "Logs");

            var logLevel = appSettings?.Logging?.LogLevel ?? "Information";

            // Ensure log directory exists
            Directory.CreateDirectory(logDirectory);

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Is(ParseLogLevel(logLevel))
                .Enrich.FromLogContext()
                .WriteTo.Console(outputTemplate: "[{Timestamp:HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}")
                .WriteTo.File(
                    path: Path.Combine(logDirectory, "bulkeditor-.log"),
                    rollingInterval: RollingInterval.Day,
                    outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} {Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            Log.Information("Application starting up");
        }

        private static Serilog.Events.LogEventLevel ParseLogLevel(string logLevel)
        {
            return logLevel?.ToLowerInvariant() switch
            {
                "verbose" => Serilog.Events.LogEventLevel.Verbose,
                "debug" => Serilog.Events.LogEventLevel.Debug,
                "information" => Serilog.Events.LogEventLevel.Information,
                "warning" => Serilog.Events.LogEventLevel.Warning,
                "error" => Serilog.Events.LogEventLevel.Error,
                "fatal" => Serilog.Events.LogEventLevel.Fatal,
                _ => Serilog.Events.LogEventLevel.Information
            };
        }
    }
}
