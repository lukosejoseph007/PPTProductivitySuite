using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTProductivitySuite
{
    public class ThemeColorManager
    {
        public class PowerPointThemeColors
        {
            public Color Background1 { get; set; }
            public Color Text1 { get; set; }
            public Color Background2 { get; set; }
            public Color Text2 { get; set; }
            public Color Accent1 { get; set; }
            public Color Accent2 { get; set; }
            public Color Accent3 { get; set; }
            public Color Accent4 { get; set; }
            public Color Accent5 { get; set; }
            public Color Accent6 { get; set; }
            public Color Hyperlink { get; set; }
            public Color FollowedHyperlink { get; set; }
        }

        public class MermaidColorMapping
        {
            public string PrimaryColor { get; set; } = "Accent1";
            public string SecondaryColor { get; set; } = "Accent2";
            public string TertiaryColor { get; set; } = "Accent3";
            public string PrimaryTextColor { get; set; } = "Text1";
            public string SecondaryTextColor { get; set; } = "Text2";
            public string PrimaryBorderColor { get; set; } = "Accent1";
            public string LineColor { get; set; } = "Text1";
            public string Background { get; set; } = "Background1";
            public string MainBkg { get; set; } = "Background2";
            public string LinkStroke { get; set; } = "Accent1";
            public string LinkFill { get; set; } = "Background1";
            public string LinkText { get; set; } = "Text1";

            // Predefined mappings
            public static readonly Dictionary<string, MermaidColorMapping> Presets = new Dictionary<string, MermaidColorMapping>
            {
                ["Corporate"] = new MermaidColorMapping
                {
                    PrimaryColor = "Accent1",
                    SecondaryColor = "Accent2",
                    TertiaryColor = "Accent3",
                    PrimaryTextColor = "Text1",
                    SecondaryTextColor = "Text2",
                    PrimaryBorderColor = "Accent1",
                    LineColor = "Text1",
                    Background = "Background1",
                    MainBkg = "Background2",
                    LinkStroke = "Accent1",
                    LinkFill = "Background1",
                    LinkText = "Text1"
                },
                ["Vibrant"] = new MermaidColorMapping
                {
                    PrimaryColor = "Accent3",
                    SecondaryColor = "Accent5",
                    TertiaryColor = "Accent1",
                    PrimaryTextColor = "Text1",
                    SecondaryTextColor = "Text2",
                    PrimaryBorderColor = "Accent3",
                    LineColor = "Accent1",
                    Background = "Background1",
                    MainBkg = "Background2",
                    LinkStroke = "Accent1",
                    LinkFill = "Background1",
                    LinkText = "Text1"
                },
                ["Monochrome"] = new MermaidColorMapping
                {
                    PrimaryColor = "Text1",
                    SecondaryColor = "Text2",
                    TertiaryColor = "Accent1",
                    PrimaryTextColor = "Text1",
                    SecondaryTextColor = "Text2",
                    PrimaryBorderColor = "Text1",
                    LineColor = "Text1",
                    Background = "Background1",
                    MainBkg = "Background2",
                    LinkStroke = "Text1",
                    LinkFill = "Background1",
                    LinkText = "Text1"
                }
            };
        }

        public static PowerPointThemeColors ExtractThemeColors(PowerPoint.Presentation presentation)
        {
            try
            {
                // Try the newer ThemeColorScheme approach first (Office 2007+)
                return ExtractModernThemeColors(presentation);
            }
            catch (Exception)
            {
                try
                {
                    // Fall back to legacy ColorScheme approach
                    return ExtractLegacyThemeColors(presentation);
                }
                catch (Exception)
                {
                    // Return default colors if both methods fail
                    return GetDefaultColors();
                }
            }
        }

        private static PowerPointThemeColors ExtractModernThemeColors(PowerPoint.Presentation presentation)
        {
            try
            {
                var themeColorScheme = presentation.SlideMaster.Theme.ThemeColorScheme;
                return new PowerPointThemeColors
                {
                    Background1 = ConvertMsoThemeColor(themeColorScheme, MsoThemeColorSchemeIndex.msoThemeLight1),
                    Text1 = ConvertMsoThemeColor(themeColorScheme, MsoThemeColorSchemeIndex.msoThemeDark1),
                    Background2 = ConvertMsoThemeColor(themeColorScheme, MsoThemeColorSchemeIndex.msoThemeLight2),
                    Text2 = ConvertMsoThemeColor(themeColorScheme, MsoThemeColorSchemeIndex.msoThemeDark2),
                    Accent1 = ConvertMsoThemeColor(themeColorScheme, MsoThemeColorSchemeIndex.msoThemeAccent1),
                    Accent2 = ConvertMsoThemeColor(themeColorScheme, MsoThemeColorSchemeIndex.msoThemeAccent2),
                    Accent3 = ConvertMsoThemeColor(themeColorScheme, MsoThemeColorSchemeIndex.msoThemeAccent3),
                    Accent4 = ConvertMsoThemeColor(themeColorScheme, MsoThemeColorSchemeIndex.msoThemeAccent4),
                    Accent5 = ConvertMsoThemeColor(themeColorScheme, MsoThemeColorSchemeIndex.msoThemeAccent5),
                    Accent6 = ConvertMsoThemeColor(themeColorScheme, MsoThemeColorSchemeIndex.msoThemeAccent6),
                    Hyperlink = ConvertMsoThemeColor(themeColorScheme, MsoThemeColorSchemeIndex.msoThemeHyperlink),
                    FollowedHyperlink = ConvertMsoThemeColor(themeColorScheme, MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink)
                };
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Modern theme color extraction failed: {ex.Message}", ex);
            }
        }

        private static PowerPointThemeColors ExtractLegacyThemeColors(PowerPoint.Presentation presentation)
        {
            try
            {
                var colorScheme = presentation.SlideMaster.ColorScheme;
                return new PowerPointThemeColors
                {
                    Background1 = ConvertToColor(colorScheme[PowerPoint.PpColorSchemeIndex.ppBackground]),
                    Text1 = ConvertToColor(colorScheme[PowerPoint.PpColorSchemeIndex.ppForeground]),
                    Background2 = ConvertToColor(colorScheme[PowerPoint.PpColorSchemeIndex.ppFill]),
                    Text2 = ConvertToColor(colorScheme[PowerPoint.PpColorSchemeIndex.ppShadow]),
                    Accent1 = ConvertToColor(colorScheme[PowerPoint.PpColorSchemeIndex.ppAccent1]),
                    Accent2 = ConvertToColor(colorScheme[PowerPoint.PpColorSchemeIndex.ppTitle]),
                    Accent3 = DarkenColor(ConvertToColor(colorScheme[PowerPoint.PpColorSchemeIndex.ppAccent1]), 0.2f),
                    Accent4 = LightenColor(ConvertToColor(colorScheme[PowerPoint.PpColorSchemeIndex.ppAccent1]), 0.2f),
                    Accent5 = DarkenColor(ConvertToColor(colorScheme[PowerPoint.PpColorSchemeIndex.ppTitle]), 0.2f),
                    Accent6 = LightenColor(ConvertToColor(colorScheme[PowerPoint.PpColorSchemeIndex.ppTitle]), 0.2f),
                    Hyperlink = Color.FromArgb(70, 120, 180),
                    FollowedHyperlink = Color.FromArgb(120, 70, 180)
                };
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Legacy theme color extraction failed: {ex.Message}", ex);
            }
        }

        private static PowerPointThemeColors GetDefaultColors()
        {
            return new PowerPointThemeColors
            {
                Background1 = Color.White,
                Text1 = Color.Black,
                Background2 = Color.FromArgb(240, 240, 240),
                Text2 = Color.FromArgb(68, 68, 68),
                Accent1 = Color.FromArgb(68, 114, 196),
                Accent2 = Color.FromArgb(237, 125, 49),
                Accent3 = Color.FromArgb(165, 165, 165),
                Accent4 = Color.FromArgb(255, 192, 0),
                Accent5 = Color.FromArgb(91, 155, 213),
                Accent6 = Color.FromArgb(112, 173, 71),
                Hyperlink = Color.FromArgb(70, 120, 180),
                FollowedHyperlink = Color.FromArgb(120, 70, 180)
            };
        }

        public static string GenerateMermaidThemeConfig(PowerPointThemeColors themeColors, MermaidColorMapping mapping)
        {
            var colorMap = new Dictionary<string, Color>
            {
                ["Background1"] = themeColors.Background1,
                ["Text1"] = themeColors.Text1,
                ["Background2"] = themeColors.Background2,
                ["Text2"] = themeColors.Text2,
                ["Accent1"] = themeColors.Accent1,
                ["Accent2"] = themeColors.Accent2,
                ["Accent3"] = themeColors.Accent3,
                ["Accent4"] = themeColors.Accent4,
                ["Accent5"] = themeColors.Accent5,
                ["Accent6"] = themeColors.Accent6,
                ["Hyperlink"] = themeColors.Hyperlink,
                ["FollowedHyperlink"] = themeColors.FollowedHyperlink
            };

            var themeConfig = $@"%%{{init: {{
    'theme': 'base',
    'themeVariables': {{
        'primaryColor': '{ColorToHex(colorMap[mapping.PrimaryColor])}',
        'primaryTextColor': '{ColorToHex(colorMap[mapping.PrimaryTextColor])}',
        'primaryBorderColor': '{ColorToHex(colorMap[mapping.PrimaryBorderColor])}',
        'lineColor': '{ColorToHex(colorMap[mapping.LineColor])}',
        'secondaryColor': '{ColorToHex(colorMap[mapping.SecondaryColor])}',
        'tertiaryColor': '{ColorToHex(colorMap[mapping.TertiaryColor])}',
        'background': '{ColorToHex(colorMap[mapping.Background])}',
        'mainBkg': '{ColorToHex(colorMap[mapping.MainBkg])}',
        'secondaryBkg': '{ColorToHex(LightenColor(colorMap[mapping.MainBkg], 0.1f))}',
        'tertiaryBkg': '{ColorToHex(LightenColor(colorMap[mapping.MainBkg], 0.2f))}',
        'primaryLabelColor': '{ColorToHex(colorMap[mapping.PrimaryTextColor])}',
        'secondaryLabelColor': '{ColorToHex(colorMap[mapping.SecondaryTextColor])}',
        'tertiaryLabelColor': '{ColorToHex(colorMap[mapping.SecondaryTextColor])}',
        'nodeBkg': '{ColorToHex(colorMap[mapping.MainBkg])}',
        'nodeTextColor': '{ColorToHex(colorMap[mapping.PrimaryTextColor])}',
        'edgeLabelBackground': '{ColorToHex(colorMap[mapping.Background])}',
        'clusterBkg': '{ColorToHex(LightenColor(colorMap[mapping.SecondaryColor], 0.8f))}',
        'clusterBorder': '{ColorToHex(colorMap[mapping.SecondaryColor])}',
        'edgeLabel': {{
            'color': '{ColorToHex(colorMap[mapping.LinkText])}',
            'background': '{ColorToHex(colorMap[mapping.LinkFill])}',
            'stroke': '{ColorToHex(colorMap[mapping.LinkStroke])}'
        }}
    }}
}}}}%%";

            return themeConfig;
        }

        private static Color ConvertMsoThemeColor(ThemeColorScheme themeColorScheme, MsoThemeColorSchemeIndex colorIndex)
        {
            try
            {
                var themeColor = themeColorScheme.Colors(colorIndex);
                var rgb = themeColor.RGB;
                return Color.FromArgb(
                    (int)(rgb & 0xFF),
                    (int)((rgb >> 8) & 0xFF),
                    (int)((rgb >> 16) & 0xFF)
                );
            }
            catch
            {
                return GetFallbackColor(colorIndex);
            }
        }

        private static Color ConvertToColor(PowerPoint.RGBColor rgbColor)
        {
            try
            {
                int rgb = (int)rgbColor.RGB;
                return Color.FromArgb(
                    rgb & 0xFF,
                    (rgb >> 8) & 0xFF,
                    (rgb >> 16) & 0xFF
                );
            }
            catch
            {
                return Color.Black;
            }
        }

        private static Color GetFallbackColor(MsoThemeColorSchemeIndex colorIndex)
        {
            switch (colorIndex)
            {
                case MsoThemeColorSchemeIndex.msoThemeLight1:
                    return Color.White;
                case MsoThemeColorSchemeIndex.msoThemeDark1:
                    return Color.Black;
                case MsoThemeColorSchemeIndex.msoThemeLight2:
                    return Color.FromArgb(240, 240, 240);
                case MsoThemeColorSchemeIndex.msoThemeDark2:
                    return Color.FromArgb(68, 68, 68);
                case MsoThemeColorSchemeIndex.msoThemeAccent1:
                    return Color.FromArgb(68, 114, 196);
                case MsoThemeColorSchemeIndex.msoThemeAccent2:
                    return Color.FromArgb(237, 125, 49);
                case MsoThemeColorSchemeIndex.msoThemeAccent3:
                    return Color.FromArgb(165, 165, 165);
                case MsoThemeColorSchemeIndex.msoThemeAccent4:
                    return Color.FromArgb(255, 192, 0);
                case MsoThemeColorSchemeIndex.msoThemeAccent5:
                    return Color.FromArgb(91, 155, 213);
                case MsoThemeColorSchemeIndex.msoThemeAccent6:
                    return Color.FromArgb(112, 173, 71);
                case MsoThemeColorSchemeIndex.msoThemeHyperlink:
                    return Color.FromArgb(70, 120, 180);
                case MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink:
                    return Color.FromArgb(120, 70, 180);
                default:
                    return Color.Gray;
            }
        }

        private static string ColorToHex(Color color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        private static Color DarkenColor(Color color, float factor)
        {
            return Color.FromArgb(
                (int)(color.R * (1 - factor)),
                (int)(color.G * (1 - factor)),
                (int)(color.B * (1 - factor))
            );
        }

        private static Color LightenColor(Color color, float factor)
        {
            return Color.FromArgb(
                Math.Min(255, (int)(color.R + (255 - color.R) * factor)),
                Math.Min(255, (int)(color.G + (255 - color.G) * factor)),
                Math.Min(255, (int)(color.B + (255 - color.B) * factor))
            );
        }
    }
}