using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using BulkEditor.Core.Configuration;
using BulkEditor.UI.Services;

namespace BulkEditor.UI.ViewModels.Settings
{
    /// <summary>
    /// ViewModel for theme settings configuration
    /// </summary>
    public partial class ThemeSettingsViewModel : ObservableObject
    {
        private readonly IThemeManager _themeManager;
        private ThemeSettings _themeSettings;

        // Base Theme Properties
        [ObservableProperty]
        private string _baseTheme = "Light";

        [ObservableProperty]
        private bool _useSystemTheme = true;

        [ObservableProperty]
        private bool _enableMaterialDesign = true;

        // Color Scheme Properties
        [ObservableProperty]
        private string _primaryColor = "#2196F3";

        [ObservableProperty]
        private string _secondaryColor = "#03DAC6";

        [ObservableProperty]
        private string _backgroundColor = "#FAFAFA";

        [ObservableProperty]
        private string _surfaceColor = "#FFFFFF";

        [ObservableProperty]
        private string _errorColor = "#F44336";

        [ObservableProperty]
        private string _warningColor = "#FF9800";

        [ObservableProperty]
        private string _successColor = "#4CAF50";

        [ObservableProperty]
        private string _infoColor = "#2196F3";

        // Typography Properties
        [ObservableProperty]
        private string _fontFamily = "Segoe UI";

        [ObservableProperty]
        private double _fontSizeScale = 1.0;

        // Layout Properties
        [ObservableProperty]
        private double _baseSpacing = 8.0;

        [ObservableProperty]
        private double _cornerRadius = 6.0;

        [ObservableProperty]
        private double _cardRadius = 12.0;

        [ObservableProperty]
        private string _density = "Normal";

        // Effects Properties
        [ObservableProperty]
        private bool _enableShadows = true;

        [ObservableProperty]
        private bool _enableAnimations = true;

        [ObservableProperty]
        private bool _enableRipple = true;

        [ObservableProperty]
        private double _animationSpeed = 1.0;

        // Advanced Properties
        [ObservableProperty]
        private bool _enableCustomPalette = false;

        [ObservableProperty]
        private string _customThemeName = "Custom";

        // Collections
        public ObservableCollection<string> AvailableThemes { get; }
        public ObservableCollection<string> AvailableFonts { get; }
        public ObservableCollection<string> AvailableDensities { get; }
        public ObservableCollection<ColorOption> PrimaryColors { get; }
        public ObservableCollection<ColorOption> SecondaryColors { get; }

        public ThemeSettingsViewModel(IThemeManager themeManager)
        {
            _themeManager = themeManager ?? throw new ArgumentNullException(nameof(themeManager));
            
            // Initialize collections
            AvailableThemes = new ObservableCollection<string> { "Light", "Dark", "Auto" };
            AvailableFonts = new ObservableCollection<string> 
            { 
                "Segoe UI", "Calibri", "Arial", "Verdana", "Tahoma", "Times New Roman", "Georgia" 
            };
            AvailableDensities = new ObservableCollection<string> { "Compact", "Normal", "Comfortable" };
            
            PrimaryColors = new ObservableCollection<ColorOption>();
            SecondaryColors = new ObservableCollection<ColorOption>();

            InitializeColorOptions();
            LoadCurrentTheme();

            // Subscribe to property changes for live preview
            PropertyChanged += OnPropertyChanged;
        }

        private void InitializeColorOptions()
        {
            var primaryColors = _themeManager.GetAvailablePrimaryColors();
            var secondaryColors = _themeManager.GetAvailableSecondaryColors();

            PrimaryColors.Clear();
            SecondaryColors.Clear();

            foreach (var color in primaryColors)
            {
                PrimaryColors.Add(new ColorOption { Name = GetColorName(color), Value = color });
            }

            foreach (var color in secondaryColors)
            {
                SecondaryColors.Add(new ColorOption { Name = GetColorName(color), Value = color });
            }
        }

        private string GetColorName(string hexColor)
        {
            return hexColor switch
            {
                "#F44336" => "Red",
                "#E91E63" => "Pink",
                "#9C27B0" => "Purple",
                "#673AB7" => "Deep Purple",
                "#3F51B5" => "Indigo",
                "#2196F3" => "Blue",
                "#03A9F4" => "Light Blue",
                "#00BCD4" => "Cyan",
                "#009688" => "Teal",
                "#4CAF50" => "Green",
                "#8BC34A" => "Light Green",
                "#CDDC39" => "Lime",
                "#FFEB3B" => "Yellow",
                "#FFC107" => "Amber",
                "#FF9800" => "Orange",
                "#FF5722" => "Deep Orange",
                "#795548" => "Brown",
                "#9E9E9E" => "Grey",
                "#607D8B" => "Blue Grey",
                _ => hexColor
            };
        }

        private void LoadCurrentTheme()
        {
            _themeSettings = _themeManager.GetCurrentTheme();
            
            BaseTheme = _themeSettings.BaseTheme;
            UseSystemTheme = _themeSettings.UseSystemTheme;
            EnableMaterialDesign = _themeSettings.EnableMaterialDesign;

            PrimaryColor = _themeSettings.Colors.PrimaryColor;
            SecondaryColor = _themeSettings.Colors.SecondaryColor;
            BackgroundColor = _themeSettings.Colors.BackgroundColor;
            SurfaceColor = _themeSettings.Colors.SurfaceColor;
            ErrorColor = _themeSettings.Colors.ErrorColor;
            WarningColor = _themeSettings.Colors.WarningColor;
            SuccessColor = _themeSettings.Colors.SuccessColor;
            InfoColor = _themeSettings.Colors.InfoColor;

            FontFamily = _themeSettings.Typography.FontFamily;
            FontSizeScale = _themeSettings.Typography.FontSizeScale;

            BaseSpacing = _themeSettings.Layout.BaseSpacing;
            CornerRadius = _themeSettings.Layout.CornerRadius;
            CardRadius = _themeSettings.Layout.CardRadius;
            Density = _themeSettings.Layout.Density;

            EnableShadows = _themeSettings.Effects.EnableShadows;
            EnableAnimations = _themeSettings.Effects.EnableAnimations;
            EnableRipple = _themeSettings.Effects.EnableRipple;
            AnimationSpeed = _themeSettings.Effects.AnimationSpeed;

            EnableCustomPalette = _themeSettings.EnableCustomPalette;
            CustomThemeName = _themeSettings.CustomThemeName;
        }

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            // Apply live preview for certain properties
            if (e.PropertyName != nameof(EnableCustomPalette) && 
                e.PropertyName != nameof(CustomThemeName))
            {
                ApplyLivePreview();
            }
        }

        private void ApplyLivePreview()
        {
            try
            {
                var previewTheme = CreateThemeFromSettings();
                _themeManager.ApplyTheme(previewTheme);
            }
            catch (Exception ex)
            {
                // Log error but don't crash the UI
                System.Diagnostics.Debug.WriteLine($"Error applying live preview: {ex.Message}");
            }
        }

        private ThemeSettings CreateThemeFromSettings()
        {
            return new ThemeSettings
            {
                BaseTheme = BaseTheme,
                UseSystemTheme = UseSystemTheme,
                EnableMaterialDesign = EnableMaterialDesign,
                Colors = new ColorScheme
                {
                    PrimaryColor = PrimaryColor,
                    SecondaryColor = SecondaryColor,
                    BackgroundColor = BackgroundColor,
                    SurfaceColor = SurfaceColor,
                    ErrorColor = ErrorColor,
                    WarningColor = WarningColor,
                    SuccessColor = SuccessColor,
                    InfoColor = InfoColor
                },
                Typography = new TypographySettings
                {
                    FontFamily = FontFamily,
                    FontSizeScale = FontSizeScale
                },
                Layout = new LayoutSettings
                {
                    BaseSpacing = BaseSpacing,
                    CornerRadius = CornerRadius,
                    CardRadius = CardRadius,
                    Density = Density
                },
                Effects = new EffectsSettings
                {
                    EnableShadows = EnableShadows,
                    EnableAnimations = EnableAnimations,
                    EnableRipple = EnableRipple,
                    AnimationSpeed = AnimationSpeed
                },
                EnableCustomPalette = EnableCustomPalette,
                CustomThemeName = CustomThemeName,
                LastModified = DateTime.UtcNow
            };
        }

        [RelayCommand]
        private void ResetToDefaults()
        {
            var defaultTheme = new ThemeSettings();
            _themeSettings = defaultTheme;
            LoadCurrentTheme();
            _themeManager.ApplyTheme(defaultTheme);
        }

        [RelayCommand]
        private void ToggleBaseTheme()
        {
            _themeManager.ToggleBaseTheme();
            LoadCurrentTheme();
        }

        [RelayCommand]
        private void ApplyColorPreset(string presetName)
        {
            switch (presetName)
            {
                case "Blue":
                    PrimaryColor = "#2196F3";
                    SecondaryColor = "#03DAC6";
                    break;
                case "Green":
                    PrimaryColor = "#4CAF50";
                    SecondaryColor = "#FF9800";
                    break;
                case "Purple":
                    PrimaryColor = "#9C27B0";
                    SecondaryColor = "#4CAF50";
                    break;
                case "Orange":
                    PrimaryColor = "#FF9800";
                    SecondaryColor = "#2196F3";
                    break;
            }
        }

        public ThemeSettings GetThemeSettings()
        {
            return CreateThemeFromSettings();
        }

        public void SaveSettings()
        {
            _themeSettings = CreateThemeFromSettings();
        }
    }

    /// <summary>
    /// Represents a color option for selection
    /// </summary>
    public class ColorOption
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }
}