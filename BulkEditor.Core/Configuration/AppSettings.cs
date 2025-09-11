using System;
using System.Collections.Generic;

namespace BulkEditor.Core.Configuration
{
    /// <summary>
    /// Main application settings configuration
    /// </summary>
    public class AppSettings
    {
        public ProcessingSettings Processing { get; set; } = new();
        public ValidationSettings Validation { get; set; } = new();
        public BackupSettings Backup { get; set; } = new();
        public LoggingSettings Logging { get; set; } = new();
        public UiSettings UI { get; set; } = new();
        public ApiSettings Api { get; set; } = new();
        public ReplacementSettings Replacement { get; set; } = new();
        public UpdateSettings Update { get; set; } = new();
        public OfflineSettings Offline { get; set; } = new();
    }

    /// <summary>
    /// Document processing settings
    /// </summary>
    public class ProcessingSettings
    {
        public int MaxConcurrentDocuments { get; set; } = 200;
        public int BatchSize { get; set; } = 50;
        public TimeSpan TimeoutPerDocument { get; set; } = TimeSpan.FromMinutes(5);
        public bool CreateBackupBeforeProcessing { get; set; } = true;
        public bool ValidateHyperlinks { get; set; } = true;
        public bool UpdateHyperlinks { get; set; } = true;
        public bool AddContentIds { get; set; } = true;
        public bool OptimizeText { get; set; } = false;
        public List<string> SupportedExtensions { get; set; } = new() { ".docx", ".docm" };
        public string LookupIdPattern { get; set; } = @"(TSRC-[^-]+-[0-9]{6}|CMS-[^-]+-[0-9]{6})";
    }

    /// <summary>
    /// Hyperlink validation settings
    /// </summary>
    public class ValidationSettings
    {
        public TimeSpan HttpTimeout { get; set; } = TimeSpan.FromSeconds(30);
        public int MaxRetryAttempts { get; set; } = 3;
        public TimeSpan RetryDelay { get; set; } = TimeSpan.FromSeconds(2);
        public string UserAgent { get; set; } = "BulkEditor/1.0";
        public bool CheckExpiredContent { get; set; } = true;
        public bool FollowRedirects { get; set; } = true;
        public List<string> SkipDomains { get; set; } = new();
        public bool AutoReplaceTitles { get; set; } = false;
        public bool ReportTitleDifferences { get; set; } = true;
        public bool ValidateTitlesOnly { get; set; } = false;
    }

    /// <summary>
    /// Backup and recovery settings
    /// </summary>
    public class BackupSettings
    {
        public string BackupDirectory { get; set; } = "Backups";
        public bool CreateTimestampedBackups { get; set; } = true;
        public int MaxBackupAge { get; set; } = 30; // days
        public bool CompressBackups { get; set; } = false;
        public bool AutoCleanupOldBackups { get; set; } = true;
    }

    /// <summary>
    /// Logging configuration settings
    /// </summary>
    public class LoggingSettings
    {
        public string LogLevel { get; set; } = "Information";
        public string LogDirectory { get; set; } = "Logs";
        public bool EnableFileLogging { get; set; } = true;
        public bool EnableConsoleLogging { get; set; } = true;
        public int MaxLogFileSizeMB { get; set; } = 10;
        public int MaxLogFiles { get; set; } = 5;
        public string LogFilePattern { get; set; } = "bulkeditor-{Date}.log";
    }

    /// <summary>
    /// User interface settings
    /// </summary>
    public class UiSettings
    {
        public string Theme { get; set; } = "Light";
        public string Language { get; set; } = "en-US";
        public bool ShowProgressDetails { get; set; } = true;
        public bool MinimizeToSystemTray { get; set; } = false;
        public bool ConfirmBeforeProcessing { get; set; } = true;
        public bool AutoSaveSettings { get; set; } = true;
        public string ConsultantEmail { get; set; } = string.Empty; // CORRECT LOCATION
        public WindowSettings Window { get; set; } = new();
        public ThemeSettings ThemeConfiguration { get; set; } = new();
    }

    /// <summary>
    /// Window settings
    /// </summary>
    public class WindowSettings
    {
        public double Width { get; set; } = 1200;
        public double Height { get; set; } = 800;
        public double Left { get; set; } = 100;
        public double Top { get; set; } = 100;
        public bool Maximized { get; set; } = false;
    }

    /// <summary>
    /// API and web service settings
    /// </summary>
    public class ApiSettings
    {
        public string BaseUrl { get; set; } = string.Empty;
        public string ApiKey { get; set; } = string.Empty;
        public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(30);
        public bool EnableCaching { get; set; } = true;
        public TimeSpan CacheExpiry { get; set; } = TimeSpan.FromHours(1);
    }
    /// <summary>
    /// Replacement processing settings
    /// </summary>
    public class ReplacementSettings
    {
        public bool EnableHyperlinkReplacement { get; set; } = false;
        public bool EnableTextReplacement { get; set; } = false;
        public List<HyperlinkReplacementRule> HyperlinkRules { get; set; } = new();
        public List<TextReplacementRule> TextRules { get; set; } = new();
        public int MaxReplacementRules { get; set; } = 50;
        public bool ValidateContentIds { get; set; } = true;
    }

    /// <summary>
    /// Rule for hyperlink replacement
    /// </summary>
    public class HyperlinkReplacementRule
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string TitleToMatch { get; set; } = string.Empty;
        public string ContentId { get; set; } = string.Empty;
        public bool IsEnabled { get; set; } = true;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }

    /// <summary>
    /// Rule for text replacement
    /// </summary>
    public class TextReplacementRule
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string SourceText { get; set; } = string.Empty;
        public string ReplacementText { get; set; } = string.Empty;
        public bool IsEnabled { get; set; } = true;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }

    /// <summary>
    /// Auto-update settings
    /// </summary>
    public class UpdateSettings
    {
        public bool AutoUpdateEnabled { get; set; } = true;
        public int CheckIntervalHours { get; set; } = 24;
        public bool InstallSecurityUpdatesAutomatically { get; set; } = true;
        public bool NotifyOnUpdatesAvailable { get; set; } = true;
        public bool CreateBackupBeforeUpdate { get; set; } = true;
        public string GitHubOwner { get; } = "ItMeDiaTech";
        public string GitHubRepository { get; } = "Bulk_Editor_WPF";
        public DateTime LastUpdateCheck { get; set; } = DateTime.MinValue;
        public bool IncludePrerelease { get; set; } = false;
    }

    /// <summary>
    /// Offline mode and network failure handling settings
    /// </summary>
    public class OfflineSettings
    {
        public bool OfflineModeEnabled { get; set; } = false;
        public bool AllowStartupWithoutNetwork { get; set; } = true;
        public bool SkipUpdateCheckWhenOffline { get; set; } = true;
        public bool ContinueProcessingWithoutValidation { get; set; } = false;
        public int NetworkTimeoutSeconds { get; set; } = 10;
        public bool ShowOfflineIndicator { get; set; } = true;
        public bool CacheLastKnownNetworkState { get; set; } = true;
    }

    /// <summary>
    /// Comprehensive theme configuration settings
    /// </summary>
    public class ThemeSettings
    {
        // Base Theme Settings
        public string BaseTheme { get; set; } = "Light"; // Light, Dark, Auto
        public bool UseSystemTheme { get; set; } = true;
        public bool EnableMaterialDesign { get; set; } = true;
        
        // Color Scheme
        public ColorScheme Colors { get; set; } = new();
        
        // Typography
        public TypographySettings Typography { get; set; } = new();
        
        // Layout and Spacing
        public LayoutSettings Layout { get; set; } = new();
        
        // Visual Effects
        public EffectsSettings Effects { get; set; } = new();
        
        // Advanced Settings
        public bool EnableCustomPalette { get; set; } = false;
        public string CustomThemeName { get; set; } = "Custom";
        public DateTime LastModified { get; set; } = DateTime.UtcNow;
    }

    /// <summary>
    /// Color scheme configuration
    /// </summary>
    public class ColorScheme
    {
        // Primary Colors
        public string PrimaryColor { get; set; } = "#2196F3"; // Material Blue
        public string PrimaryVariant { get; set; } = "#1976D2";
        public string PrimaryLight { get; set; } = "#BBDEFB";
        
        // Secondary Colors
        public string SecondaryColor { get; set; } = "#03DAC6"; // Material Teal
        public string SecondaryVariant { get; set; } = "#018786";
        public string SecondaryLight { get; set; } = "#B2DFDB";
        
        // Surface Colors
        public string SurfaceColor { get; set; } = "#FFFFFF";
        public string BackgroundColor { get; set; } = "#FAFAFA";
        public string CardColor { get; set; } = "#FFFFFF";
        public string DialogColor { get; set; } = "#FFFFFF";
        
        // Text Colors
        public string OnPrimaryColor { get; set; } = "#FFFFFF";
        public string OnSecondaryColor { get; set; } = "#000000";
        public string OnSurfaceColor { get; set; } = "#212121";
        public string OnBackgroundColor { get; set; } = "#212121";
        
        // State Colors
        public string ErrorColor { get; set; } = "#F44336";
        public string WarningColor { get; set; } = "#FF9800";
        public string SuccessColor { get; set; } = "#4CAF50";
        public string InfoColor { get; set; } = "#2196F3";
        
        // Neutral Colors
        public string DisabledColor { get; set; } = "#BDBDBD";
        public string BorderColor { get; set; } = "#E0E0E0";
        public string DividerColor { get; set; } = "#EEEEEE";
        public string HoverColor { get; set; } = "#F5F5F5";
    }

    /// <summary>
    /// Typography settings
    /// </summary>
    public class TypographySettings
    {
        public string FontFamily { get; set; } = "Segoe UI";
        public string HeadlineFontFamily { get; set; } = "Segoe UI";
        public double FontSizeScale { get; set; } = 1.0; // Scale factor for all fonts
        
        // Font Sizes
        public double HeadlineSize { get; set; } = 24;
        public double TitleSize { get; set; } = 20;
        public double SubtitleSize { get; set; } = 16;
        public double BodySize { get; set; } = 14;
        public double CaptionSize { get; set; } = 12;
        
        // Font Weights
        public string HeadlineWeight { get; set; } = "SemiBold";
        public string TitleWeight { get; set; } = "Medium";
        public string SubtitleWeight { get; set; } = "Medium";
        public string BodyWeight { get; set; } = "Normal";
        public string CaptionWeight { get; set; } = "Normal";
    }

    /// <summary>
    /// Layout and spacing settings
    /// </summary>
    public class LayoutSettings
    {
        // Base spacing unit (Material Design recommends 8dp)
        public double BaseSpacing { get; set; } = 8;
        
        // Spacing multipliers
        public double SmallSpacing { get; set; } = 0.5; // 4dp
        public double MediumSpacing { get; set; } = 1.0; // 8dp
        public double LargeSpacing { get; set; } = 2.0; // 16dp
        public double ExtraLargeSpacing { get; set; } = 3.0; // 24dp
        
        // Corner radius
        public double CornerRadius { get; set; } = 6;
        public double CardRadius { get; set; } = 12;
        public double ButtonRadius { get; set; } = 6;
        
        // Padding and margins
        public double ContentPadding { get; set; } = 16;
        public double CardPadding { get; set; } = 24;
        public double ButtonPadding { get; set; } = 16;
        
        // Layout density
        public string Density { get; set; } = "Normal"; // Compact, Normal, Comfortable
    }

    /// <summary>
    /// Visual effects settings
    /// </summary>
    public class EffectsSettings
    {
        // Shadows and elevation
        public bool EnableShadows { get; set; } = true;
        public double ShadowOpacity { get; set; } = 0.1;
        public double ShadowBlurRadius { get; set; } = 8;
        public double ShadowDepth { get; set; } = 2;
        
        // Animations
        public bool EnableAnimations { get; set; } = true;
        public double AnimationSpeed { get; set; } = 1.0; // Speed multiplier
        public string AnimationEasing { get; set; } = "QuadraticEase";
        
        // Ripple effects
        public bool EnableRipple { get; set; } = true;
        public double RippleOpacity { get; set; } = 0.1;
        
        // Transitions
        public bool EnableTransitions { get; set; } = true;
        public int TransitionDuration { get; set; } = 200; // milliseconds
        
        // Visual enhancements
        public bool EnableBlur { get; set; } = false;
        public double BlurRadius { get; set; } = 5;
        public bool EnableGradients { get; set; } = true;
    }
}