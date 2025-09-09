using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using BulkEditor.Core.Interfaces;

namespace BulkEditor.UI.Themes
{
    /// <summary>
    /// Manages application themes and provides theme switching functionality
    /// </summary>
    public class ThemeManager
    {
        private readonly ILoggingService _logger;
        private const string THEME_RESOURCE_PATH = "Themes/";
        
        public ThemeManager(ILoggingService logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// Available theme names
        /// </summary>
        public static readonly Dictionary<string, string> AvailableThemes = new()
        {
            { "Light", "Light Theme" },
            { "Dark", "Dark Theme" },
            { "LightBlue", "Light Blue Theme" },
            { "Green", "Green Theme" },
            { "Purple", "Purple Theme" },
            { "Pink", "Pink Theme" }
        };

        /// <summary>
        /// Gets the current theme name
        /// </summary>
        public string CurrentTheme { get; private set; } = "Light";

        /// <summary>
        /// Event raised when theme changes
        /// </summary>
        public event EventHandler<ThemeChangedEventArgs>? ThemeChanged;

        /// <summary>
        /// Applies a theme to the application
        /// </summary>
        public bool ApplyTheme(string themeName)
        {
            try
            {
                if (!AvailableThemes.ContainsKey(themeName))
                {
                    _logger.LogWarning("Theme '{Theme}' not found, using Light theme", themeName);
                    themeName = "Light";
                }

                // Remove existing theme resources
                ClearCurrentTheme();

                // Load new theme resource dictionary
                var themeUri = new Uri($"pack://application:,,,/BulkEditor.UI;component/{THEME_RESOURCE_PATH}{themeName}Theme.xaml");
                var themeDict = new ResourceDictionary { Source = themeUri };

                // Add to application resources
                System.Windows.Application.Current.Resources.MergedDictionaries.Add(themeDict);

                var previousTheme = CurrentTheme;
                CurrentTheme = themeName;

                _logger.LogInformation("Theme changed from '{PreviousTheme}' to '{NewTheme}'", previousTheme, themeName);
                
                // Raise theme changed event
                ThemeChanged?.Invoke(this, new ThemeChangedEventArgs(previousTheme, CurrentTheme));

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to apply theme '{Theme}'", themeName);
                return false;
            }
        }

        /// <summary>
        /// Gets a themed brush by key
        /// </summary>
        public SolidColorBrush? GetThemedBrush(string resourceKey)
        {
            try
            {
                if (System.Windows.Application.Current.Resources.Contains(resourceKey))
                {
                    return System.Windows.Application.Current.Resources[resourceKey] as SolidColorBrush;
                }
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to get themed brush '{ResourceKey}'", resourceKey);
                return null;
            }
        }

        /// <summary>
        /// Gets a themed color by key
        /// </summary>
        public Color? GetThemedColor(string resourceKey)
        {
            try
            {
                if (System.Windows.Application.Current.Resources.Contains(resourceKey))
                {
                    if (System.Windows.Application.Current.Resources[resourceKey] is Color color)
                        return color;
                    if (System.Windows.Application.Current.Resources[resourceKey] is SolidColorBrush brush)
                        return brush.Color;
                }
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to get themed color '{ResourceKey}'", resourceKey);
                return null;
            }
        }

        /// <summary>
        /// Initializes the theme system with default theme
        /// </summary>
        public void Initialize(string defaultTheme = "Light")
        {
            try
            {
                _logger.LogInformation("Initializing theme system with default theme: {DefaultTheme}", defaultTheme);
                ApplyTheme(defaultTheme);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to initialize theme system");
                // Fallback to basic theme
                ApplyTheme("Light");
            }
        }

        /// <summary>
        /// Clears current theme resources
        /// </summary>
        private void ClearCurrentTheme()
        {
            try
            {
                // Find and remove theme resource dictionaries
                var themeResources = System.Windows.Application.Current.Resources.MergedDictionaries
                    .Where(dict => dict.Source?.OriginalString?.Contains(THEME_RESOURCE_PATH) == true)
                    .ToList();

                foreach (var resource in themeResources)
                {
                    System.Windows.Application.Current.Resources.MergedDictionaries.Remove(resource);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error clearing current theme resources");
            }
        }

        /// <summary>
        /// Gets system accent color if available
        /// </summary>
        public static Color GetSystemAccentColor()
        {
            try
            {
                // Try to get Windows accent color
                var accentColor = SystemParameters.WindowGlassColor;
                return accentColor;
            }
            catch
            {
                // Fallback to default blue
                return Color.FromRgb(0x21, 0x96, 0xF3);
            }
        }

        /// <summary>
        /// Determines if the system is using dark mode
        /// </summary>
        public static bool IsSystemDarkMode()
        {
            try
            {
                // Check registry for Windows theme preference
                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                var registryValueObject = key?.GetValue("AppsUseLightTheme");
                if (registryValueObject == null)
                    return false;
                
                var registryValue = (int)registryValueObject;
                return registryValue == 0;
            }
            catch
            {
                return false;
            }
        }
    }

    /// <summary>
    /// Event args for theme change events
    /// </summary>
    public class ThemeChangedEventArgs : EventArgs
    {
        public string PreviousTheme { get; }
        public string NewTheme { get; }

        public ThemeChangedEventArgs(string previousTheme, string newTheme)
        {
            PreviousTheme = previousTheme;
            NewTheme = newTheme;
        }
    }
}