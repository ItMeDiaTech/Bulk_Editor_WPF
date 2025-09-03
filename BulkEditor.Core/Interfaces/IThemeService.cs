using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace BulkEditor.Core.Interfaces
{
    /// <summary>
    /// Interface for theme management and switching
    /// </summary>
    public interface IThemeService
    {
        /// <summary>
        /// Gets the current active theme
        /// </summary>
        string CurrentTheme { get; }

        /// <summary>
        /// Gets available themes
        /// </summary>
        IEnumerable<string> AvailableThemes { get; }

        /// <summary>
        /// Applies a theme to the application
        /// </summary>
        Task ApplyThemeAsync(string themeName);

        /// <summary>
        /// Event raised when theme changes
        /// </summary>
        event EventHandler<ThemeChangedEventArgs>? ThemeChanged;

        /// <summary>
        /// Gets theme-specific resource dictionary path
        /// </summary>
        string GetThemeResourcePath(string themeName);

        /// <summary>
        /// Validates if theme name is supported
        /// </summary>
        bool IsThemeSupported(string themeName);
    }

    /// <summary>
    /// Event arguments for theme change notifications
    /// </summary>
    public class ThemeChangedEventArgs : EventArgs
    {
        public string PreviousTheme { get; set; } = string.Empty;
        public string NewTheme { get; set; } = string.Empty;
        public DateTime ChangedAt { get; set; } = DateTime.UtcNow;
    }

    /// <summary>
    /// Available application themes
    /// </summary>
    public static class AppThemes
    {
        public const string Light = "Light";
        public const string Dark = "Dark";
        public const string Auto = "Auto"; // Follow system theme

        public static readonly string[] All = { Light, Dark, Auto };
    }
}