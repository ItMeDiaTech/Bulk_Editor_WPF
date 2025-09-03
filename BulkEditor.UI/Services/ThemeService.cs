using BulkEditor.Core.Configuration;
using BulkEditor.Core.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace BulkEditor.UI.Services
{
    /// <summary>
    /// Implementation of theme management service for WPF application
    /// </summary>
    public class ThemeService : IThemeService
    {
        private readonly ILoggingService _logger;
        private readonly AppSettings _appSettings;
        private string _currentTheme;

        public event EventHandler<ThemeChangedEventArgs>? ThemeChanged;

        public string CurrentTheme => _currentTheme;

        public IEnumerable<string> AvailableThemes => AppThemes.All;

        public ThemeService(ILoggingService logger, AppSettings appSettings)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _appSettings = appSettings ?? throw new ArgumentNullException(nameof(appSettings));
            _currentTheme = _appSettings.UI.Theme;
        }

        public async Task ApplyThemeAsync(string themeName)
        {
            try
            {
                if (string.IsNullOrEmpty(themeName))
                    throw new ArgumentException("Theme name cannot be null or empty", nameof(themeName));

                if (!IsThemeSupported(themeName))
                    throw new ArgumentException($"Unsupported theme: {themeName}", nameof(themeName));

                var previousTheme = _currentTheme;

                // Apply theme to WPF application resources
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    ApplyThemeResources(themeName);
                });

                // Update current theme
                _currentTheme = themeName;
                _appSettings.UI.Theme = themeName;

                // Raise theme changed event
                ThemeChanged?.Invoke(this, new ThemeChangedEventArgs
                {
                    PreviousTheme = previousTheme,
                    NewTheme = themeName
                });

                _logger.LogInformation("Theme changed from '{PreviousTheme}' to '{NewTheme}'", previousTheme, themeName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error applying theme: {ThemeName}", themeName);
                throw;
            }
        }

        public string GetThemeResourcePath(string themeName)
        {
            return themeName.ToLowerInvariant() switch
            {
                "light" => "pack://application:,,,/BulkEditor.UI;component/Themes/LightTheme.xaml",
                "dark" => "pack://application:,,,/BulkEditor.UI;component/Themes/DarkTheme.xaml",
                "auto" => GetSystemTheme() == "Dark"
                    ? "pack://application:,,,/BulkEditor.UI;component/Themes/DarkTheme.xaml"
                    : "pack://application:,,,/BulkEditor.UI;component/Themes/LightTheme.xaml",
                _ => "pack://application:,,,/BulkEditor.UI;component/Themes/LightTheme.xaml"
            };
        }

        public bool IsThemeSupported(string themeName)
        {
            return AvailableThemes.Contains(themeName, StringComparer.OrdinalIgnoreCase);
        }

        private void ApplyThemeResources(string themeName)
        {
            try
            {
                var app = System.Windows.Application.Current;
                if (app?.Resources == null)
                    return;

                // Remove existing theme resources
                var existingThemeResources = app.Resources.MergedDictionaries
                    .Where(d => d.Source?.ToString().Contains("/Themes/") == true)
                    .ToList();

                foreach (var resource in existingThemeResources)
                {
                    app.Resources.MergedDictionaries.Remove(resource);
                }

                // Load new theme resource dictionary
                var themeResourcePath = GetThemeResourcePath(themeName);
                var themeResourceDict = new ResourceDictionary
                {
                    Source = new Uri(themeResourcePath)
                };

                // Add new theme resources
                app.Resources.MergedDictionaries.Add(themeResourceDict);

                _logger.LogDebug("Applied theme resources for: {ThemeName}", themeName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error applying theme resources for: {ThemeName}", themeName);
                throw;
            }
        }

        private string GetSystemTheme()
        {
            try
            {
                // Check Windows system theme preference
                var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                var value = key?.GetValue("AppsUseLightTheme");

                if (value is int lightTheme)
                {
                    return lightTheme == 0 ? "Dark" : "Light";
                }

                return "Light"; // Default fallback
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Could not detect system theme, defaulting to Light: {Error}", ex.Message);
                return "Light";
            }
        }
    }
}