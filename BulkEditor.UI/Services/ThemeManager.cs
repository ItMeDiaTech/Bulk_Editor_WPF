using System;
using System.Windows;
using System.Windows.Media;
using BulkEditor.Core.Configuration;
using BulkEditor.Core.Interfaces;
using MaterialDesignThemes.Wpf;
using MaterialDesignColors;

namespace BulkEditor.UI.Services
{
    /// <summary>
    /// Service for managing application themes and Material Design configuration
    /// </summary>
    public interface IThemeManager
    {
        /// <summary>
        /// Apply theme settings to the application
        /// </summary>
        void ApplyTheme(ThemeSettings themeSettings);

        /// <summary>
        /// Get current theme settings
        /// </summary>
        ThemeSettings GetCurrentTheme();

        /// <summary>
        /// Switch between light and dark themes
        /// </summary>
        void ToggleBaseTheme();

        /// <summary>
        /// Apply custom color scheme
        /// </summary>
        void ApplyColorScheme(ColorScheme colorScheme);

        /// <summary>
        /// Get available Material Design colors
        /// </summary>
        string[] GetAvailablePrimaryColors();

        /// <summary>
        /// Get available Material Design secondary colors
        /// </summary>
        string[] GetAvailableSecondaryColors();

        /// <summary>
        /// Event raised when theme changes
        /// </summary>
        event EventHandler<BulkEditor.Core.Interfaces.ThemeChangedEventArgs> ThemeChanged;
    }

    /// <summary>
    /// Theme manager implementation
    /// </summary>
    public class ThemeManager : IThemeManager
    {
        private readonly PaletteHelper _paletteHelper;
        private ThemeSettings _currentTheme;

        public event EventHandler<BulkEditor.Core.Interfaces.ThemeChangedEventArgs> ThemeChanged;

        public ThemeManager()
        {
            _paletteHelper = new PaletteHelper();
            _currentTheme = new ThemeSettings();
        }

        public void ApplyTheme(ThemeSettings themeSettings)
        {
            if (themeSettings == null)
                throw new ArgumentNullException(nameof(themeSettings));

            _currentTheme = themeSettings;

            try
            {
                // Get current theme
                var theme = _paletteHelper.GetTheme();

                // Apply base theme (Light/Dark)
                var baseTheme = themeSettings.BaseTheme.Equals("Dark", StringComparison.OrdinalIgnoreCase)
                    ? BaseTheme.Dark
                    : BaseTheme.Light;

                theme.SetBaseTheme(baseTheme);

                // Apply primary color
                if (ColorHelper.TryParseColor(themeSettings.Colors.PrimaryColor, out var primaryColor))
                {
                    theme.SetPrimaryColor(primaryColor);
                }

                // Apply secondary color
                if (ColorHelper.TryParseColor(themeSettings.Colors.SecondaryColor, out var secondaryColor))
                {
                    theme.SetSecondaryColor(secondaryColor);
                }

                // Set the theme
                _paletteHelper.SetTheme(theme);

                // Apply custom color scheme to application resources
                ApplyCustomColors(themeSettings.Colors);

                // Apply typography settings
                ApplyTypography(themeSettings.Typography);

                // Apply layout settings
                ApplyLayout(themeSettings.Layout);

                // Apply effects settings
                ApplyEffects(themeSettings.Effects);

                // Raise theme changed event
                ThemeChanged?.Invoke(this, new BulkEditor.Core.Interfaces.ThemeChangedEventArgs
                {
                    PreviousTheme = _currentTheme?.BaseTheme ?? "Light",
                    NewTheme = themeSettings.BaseTheme
                });
            }
            catch (Exception ex)
            {
                // Log error and fallback to default theme
                System.Diagnostics.Debug.WriteLine($"Error applying theme: {ex.Message}");
                ApplyDefaultTheme();
            }
        }

        public ThemeSettings GetCurrentTheme()
        {
            return _currentTheme ?? new ThemeSettings();
        }

        public void ToggleBaseTheme()
        {
            var currentTheme = GetCurrentTheme();
            currentTheme.BaseTheme = currentTheme.BaseTheme.Equals("Light", StringComparison.OrdinalIgnoreCase)
                ? "Dark"
                : "Light";

            ApplyTheme(currentTheme);
        }

        public void ApplyColorScheme(ColorScheme colorScheme)
        {
            if (colorScheme == null)
                throw new ArgumentNullException(nameof(colorScheme));

            var currentTheme = GetCurrentTheme();
            currentTheme.Colors = colorScheme;
            ApplyTheme(currentTheme);
        }

        public string[] GetAvailablePrimaryColors()
        {
            return new[]
            {
                "#F44336", // Red
                "#E91E63", // Pink
                "#9C27B0", // Purple
                "#673AB7", // Deep Purple
                "#3F51B5", // Indigo
                "#2196F3", // Blue
                "#03A9F4", // Light Blue
                "#00BCD4", // Cyan
                "#009688", // Teal
                "#4CAF50", // Green
                "#8BC34A", // Light Green
                "#CDDC39", // Lime
                "#FFEB3B", // Yellow
                "#FFC107", // Amber
                "#FF9800", // Orange
                "#FF5722", // Deep Orange
                "#795548", // Brown
                "#9E9E9E", // Grey
                "#607D8B"  // Blue Grey
            };
        }

        public string[] GetAvailableSecondaryColors()
        {
            return new[]
            {
                "#FF5722", // Deep Orange
                "#FF9800", // Orange
                "#FFC107", // Amber
                "#FFEB3B", // Yellow
                "#CDDC39", // Lime
                "#8BC34A", // Light Green
                "#4CAF50", // Green
                "#009688", // Teal
                "#00BCD4", // Cyan
                "#03A9F4", // Light Blue
                "#2196F3", // Blue
                "#3F51B5", // Indigo
                "#673AB7", // Deep Purple
                "#9C27B0", // Purple
                "#E91E63", // Pink
                "#F44336"  // Red
            };
        }

        private void ApplyCustomColors(ColorScheme colors)
        {
            var app = System.Windows.Application.Current;
            if (app?.Resources == null) return;

            // Update application-level color resources
            UpdateResourceSafely(app.Resources, "PrimaryBrush", CreateBrushFromHex(colors.PrimaryColor));
            UpdateResourceSafely(app.Resources, "PrimaryVariantBrush", CreateBrushFromHex(colors.PrimaryVariant));
            UpdateResourceSafely(app.Resources, "SecondaryBrush", CreateBrushFromHex(colors.SecondaryColor));
            UpdateResourceSafely(app.Resources, "SecondaryVariantBrush", CreateBrushFromHex(colors.SecondaryVariant));

            UpdateResourceSafely(app.Resources, "SurfaceBrush", CreateBrushFromHex(colors.SurfaceColor));
            UpdateResourceSafely(app.Resources, "BackgroundBrush", CreateBrushFromHex(colors.BackgroundColor));
            UpdateResourceSafely(app.Resources, "CardBrush", CreateBrushFromHex(colors.CardColor));

            UpdateResourceSafely(app.Resources, "OnPrimaryBrush", CreateBrushFromHex(colors.OnPrimaryColor));
            UpdateResourceSafely(app.Resources, "OnSecondaryBrush", CreateBrushFromHex(colors.OnSecondaryColor));
            UpdateResourceSafely(app.Resources, "OnSurfaceBrush", CreateBrushFromHex(colors.OnSurfaceColor));
            UpdateResourceSafely(app.Resources, "OnBackgroundBrush", CreateBrushFromHex(colors.OnBackgroundColor));

            UpdateResourceSafely(app.Resources, "ErrorBrush", CreateBrushFromHex(colors.ErrorColor));
            UpdateResourceSafely(app.Resources, "WarningBrush", CreateBrushFromHex(colors.WarningColor));
            UpdateResourceSafely(app.Resources, "SuccessBrush", CreateBrushFromHex(colors.SuccessColor));
            UpdateResourceSafely(app.Resources, "InfoBrush", CreateBrushFromHex(colors.InfoColor));

            UpdateResourceSafely(app.Resources, "DisabledBrush", CreateBrushFromHex(colors.DisabledColor));
            UpdateResourceSafely(app.Resources, "BorderBrush", CreateBrushFromHex(colors.BorderColor));
            UpdateResourceSafely(app.Resources, "DividerBrush", CreateBrushFromHex(colors.DividerColor));
            UpdateResourceSafely(app.Resources, "HoverBrush", CreateBrushFromHex(colors.HoverColor));
        }

        private void ApplyTypography(TypographySettings typography)
        {
            var app = System.Windows.Application.Current;
            if (app?.Resources == null) return;

            UpdateResourceSafely(app.Resources, "AppFontFamily", new FontFamily(typography.FontFamily));
            UpdateResourceSafely(app.Resources, "HeadlineFontFamily", new FontFamily(typography.HeadlineFontFamily));

            UpdateResourceSafely(app.Resources, "HeadlineFontSize", typography.HeadlineSize * typography.FontSizeScale);
            UpdateResourceSafely(app.Resources, "TitleFontSize", typography.TitleSize * typography.FontSizeScale);
            UpdateResourceSafely(app.Resources, "SubtitleFontSize", typography.SubtitleSize * typography.FontSizeScale);
            UpdateResourceSafely(app.Resources, "BodyFontSize", typography.BodySize * typography.FontSizeScale);
            UpdateResourceSafely(app.Resources, "CaptionFontSize", typography.CaptionSize * typography.FontSizeScale);

            UpdateResourceSafely(app.Resources, "HeadlineFontWeight", ParseFontWeight(typography.HeadlineWeight));
            UpdateResourceSafely(app.Resources, "TitleFontWeight", ParseFontWeight(typography.TitleWeight));
            UpdateResourceSafely(app.Resources, "SubtitleFontWeight", ParseFontWeight(typography.SubtitleWeight));
            UpdateResourceSafely(app.Resources, "BodyFontWeight", ParseFontWeight(typography.BodyWeight));
            UpdateResourceSafely(app.Resources, "CaptionFontWeight", ParseFontWeight(typography.CaptionWeight));
        }

        private void ApplyLayout(LayoutSettings layout)
        {
            var app = System.Windows.Application.Current;
            if (app?.Resources == null) return;

            // Numeric spacing values for Width, Height, and other numeric properties
            UpdateResourceSafely(app.Resources, "BaseSpacingValue", layout.BaseSpacing);
            UpdateResourceSafely(app.Resources, "SmallSpacingValue", layout.BaseSpacing * layout.SmallSpacing);
            UpdateResourceSafely(app.Resources, "MediumSpacingValue", layout.BaseSpacing * layout.MediumSpacing);
            UpdateResourceSafely(app.Resources, "LargeSpacingValue", layout.BaseSpacing * layout.LargeSpacing);
            UpdateResourceSafely(app.Resources, "ExtraLargeSpacingValue", layout.BaseSpacing * layout.ExtraLargeSpacing);

            // Thickness spacing values for Margin properties
            UpdateResourceSafely(app.Resources, "BaseSpacing", new Thickness(layout.BaseSpacing));
            UpdateResourceSafely(app.Resources, "SmallSpacing", new Thickness(layout.BaseSpacing * layout.SmallSpacing));
            UpdateResourceSafely(app.Resources, "MediumSpacing", new Thickness(layout.BaseSpacing * layout.MediumSpacing));
            UpdateResourceSafely(app.Resources, "LargeSpacing", new Thickness(layout.BaseSpacing * layout.LargeSpacing));
            UpdateResourceSafely(app.Resources, "ExtraLargeSpacing", new Thickness(layout.BaseSpacing * layout.ExtraLargeSpacing));

            UpdateResourceSafely(app.Resources, "CornerRadius", new CornerRadius(layout.CornerRadius));
            UpdateResourceSafely(app.Resources, "CardCornerRadius", new CornerRadius(layout.CardRadius));
            UpdateResourceSafely(app.Resources, "ButtonCornerRadius", new CornerRadius(layout.ButtonRadius));

            UpdateResourceSafely(app.Resources, "ContentPadding", new Thickness(layout.ContentPadding));
            UpdateResourceSafely(app.Resources, "CardPadding", new Thickness(layout.CardPadding));
            UpdateResourceSafely(app.Resources, "ButtonPadding", new Thickness(layout.ButtonPadding));
        }

        private void ApplyEffects(EffectsSettings effects)
        {
            var app = System.Windows.Application.Current;
            if (app?.Resources == null) return;

            UpdateResourceSafely(app.Resources, "EnableShadows", effects.EnableShadows);
            UpdateResourceSafely(app.Resources, "ShadowOpacity", effects.ShadowOpacity);
            UpdateResourceSafely(app.Resources, "ShadowBlurRadius", effects.ShadowBlurRadius);
            UpdateResourceSafely(app.Resources, "ShadowDepth", effects.ShadowDepth);

            UpdateResourceSafely(app.Resources, "EnableAnimations", effects.EnableAnimations);
            UpdateResourceSafely(app.Resources, "AnimationSpeed", effects.AnimationSpeed);
            UpdateResourceSafely(app.Resources, "TransitionDuration", TimeSpan.FromMilliseconds(effects.TransitionDuration));

            UpdateResourceSafely(app.Resources, "EnableRipple", effects.EnableRipple);
            UpdateResourceSafely(app.Resources, "RippleOpacity", effects.RippleOpacity);
        }

        private void ApplyDefaultTheme()
        {
            ApplyTheme(new ThemeSettings());
        }

        private static Brush CreateBrushFromHex(string hexColor)
        {
            if (ColorHelper.TryParseColor(hexColor, out var color))
            {
                return new SolidColorBrush(color);
            }
            return Brushes.Transparent;
        }

        private static void UpdateResourceSafely(ResourceDictionary resources, string key, object value)
        {
            try
            {
                if (resources.Contains(key))
                    resources[key] = value;
                else
                    resources.Add(key, value);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating resource '{key}' with value '{value}' (Type: {value?.GetType().Name}): {ex.Message}");
                // Log the stack trace to help identify where the error is coming from
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }

        private static FontWeight ParseFontWeight(string weight)
        {
            return weight.ToLower() switch
            {
                "thin" => FontWeights.Thin,
                "extralight" => FontWeights.ExtraLight,
                "light" => FontWeights.Light,
                "normal" => FontWeights.Normal,
                "medium" => FontWeights.Medium,
                "semibold" => FontWeights.SemiBold,
                "bold" => FontWeights.Bold,
                "extrabold" => FontWeights.ExtraBold,
                "black" => FontWeights.Black,
                _ => FontWeights.Normal
            };
        }
    }


    /// <summary>
    /// Helper class for color parsing and manipulation
    /// </summary>
    public static class ColorHelper
    {
        public static bool TryParseColor(string hexColor, out Color color)
        {
            color = Colors.Transparent;

            if (string.IsNullOrWhiteSpace(hexColor))
                return false;

            try
            {
                // Remove # if present
                var hex = hexColor.TrimStart('#');

                // Parse the color
                if (hex.Length == 6)
                {
                    color = Color.FromRgb(
                        Convert.ToByte(hex.Substring(0, 2), 16),
                        Convert.ToByte(hex.Substring(2, 2), 16),
                        Convert.ToByte(hex.Substring(4, 2), 16));
                    return true;
                }
                else if (hex.Length == 8)
                {
                    color = Color.FromArgb(
                        Convert.ToByte(hex.Substring(0, 2), 16),
                        Convert.ToByte(hex.Substring(2, 2), 16),
                        Convert.ToByte(hex.Substring(4, 2), 16),
                        Convert.ToByte(hex.Substring(6, 2), 16));
                    return true;
                }
            }
            catch
            {
                // Ignore parsing errors
            }

            return false;
        }

        public static string ToHexString(Color color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        public static string ToHexStringWithAlpha(Color color)
        {
            return $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
        }
    }
}