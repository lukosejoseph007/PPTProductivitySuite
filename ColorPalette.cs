using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Xml.Serialization;

namespace PPTProductivitySuite
{
    [Serializable]
    public class ColorPalette
    {
        public string Name { get; set; }
        public Color PrimaryColor { get; set; }
        public Color SecondaryColor { get; set; }
        public Color TertiaryColor { get; set; }
        public Color QuaternaryColor { get; set; }
        public Color PrimaryTextColor { get; set; }
        public Color SecondaryTextColor { get; set; }
        public Color BackgroundColor { get; set; }
        public Color BorderColor { get; set; }
        public Color LineColor { get; set; }
        public Color AccentColor { get; set; }

        public ColorPalette()
        {
            // Default constructor for serialization
        }

        public ColorPalette(string name)
        {
            Name = name;
            SetDefaultColors();
        }

        private void SetDefaultColors()
        {
            PrimaryColor = Color.FromArgb(68, 114, 196);
            SecondaryColor = Color.FromArgb(237, 125, 49);
            TertiaryColor = Color.FromArgb(165, 165, 165);
            QuaternaryColor = Color.FromArgb(255, 192, 0);
            PrimaryTextColor = Color.Black;
            SecondaryTextColor = Color.FromArgb(68, 68, 68);
            BackgroundColor = Color.White;
            BorderColor = Color.FromArgb(68, 114, 196);
            LineColor = Color.Black;
            AccentColor = Color.FromArgb(91, 155, 213);
        }

        // Built-in presets
        public static Dictionary<string, ColorPalette> GetBuiltInPresets()
        {
            return new Dictionary<string, ColorPalette>
            {
                ["Corporate Blue"] = new ColorPalette("Corporate Blue")
                {
                    PrimaryColor = Color.FromArgb(68, 114, 196),
                    SecondaryColor = Color.FromArgb(91, 155, 213),
                    TertiaryColor = Color.FromArgb(165, 165, 165),
                    QuaternaryColor = Color.FromArgb(112, 173, 71),
                    PrimaryTextColor = Color.Black,
                    SecondaryTextColor = Color.FromArgb(68, 68, 68),
                    BackgroundColor = Color.White,
                    BorderColor = Color.FromArgb(68, 114, 196),
                    LineColor = Color.FromArgb(68, 68, 68),
                    AccentColor = Color.FromArgb(237, 125, 49)
                },
                ["Vibrant"] = new ColorPalette("Vibrant")
                {
                    PrimaryColor = Color.FromArgb(255, 87, 51),
                    SecondaryColor = Color.FromArgb(25, 181, 254),
                    TertiaryColor = Color.FromArgb(255, 206, 84),
                    QuaternaryColor = Color.FromArgb(129, 199, 132),
                    PrimaryTextColor = Color.FromArgb(33, 33, 33),
                    SecondaryTextColor = Color.FromArgb(117, 117, 117),
                    BackgroundColor = Color.White,
                    BorderColor = Color.FromArgb(255, 87, 51),
                    LineColor = Color.FromArgb(66, 66, 66),
                    AccentColor = Color.FromArgb(156, 39, 176)
                },
                ["Dark Professional"] = new ColorPalette("Dark Professional")
                {
                    PrimaryColor = Color.FromArgb(52, 73, 94),
                    SecondaryColor = Color.FromArgb(149, 165, 166),
                    TertiaryColor = Color.FromArgb(52, 152, 219),
                    QuaternaryColor = Color.FromArgb(39, 174, 96),
                    PrimaryTextColor = Color.White,
                    SecondaryTextColor = Color.FromArgb(189, 195, 199),
                    BackgroundColor = Color.FromArgb(44, 62, 80),
                    BorderColor = Color.FromArgb(149, 165, 166),
                    LineColor = Color.FromArgb(189, 195, 199),
                    AccentColor = Color.FromArgb(231, 76, 60)
                }
            };
        }

        public string GenerateMermaidThemeConfig()
        {
            var config = $@"%%{{init: {{
    'theme': 'base',
    'themeVariables': {{
        'primaryColor': '{ColorToHex(PrimaryColor)}',
        'primaryTextColor': '{ColorToHex(PrimaryTextColor)}',
        'primaryBorderColor': '{ColorToHex(BorderColor)}',
        'lineColor': '{ColorToHex(LineColor)}',
        'secondaryColor': '{ColorToHex(SecondaryColor)}',
        'tertiaryColor': '{ColorToHex(TertiaryColor)}',
        'background': '{ColorToHex(BackgroundColor)}',
        'mainBkg': '{ColorToHex(BackgroundColor)}',
        'secondaryBkg': '{ColorToHex(LightenColor(BackgroundColor, 0.1f))}',
        'tertiaryBkg': '{ColorToHex(LightenColor(BackgroundColor, 0.2f))}',
        'primaryLabelColor': '{ColorToHex(PrimaryTextColor)}',
        'secondaryLabelColor': '{ColorToHex(SecondaryTextColor)}',
        'tertiaryLabelColor': '{ColorToHex(SecondaryTextColor)}',
        'nodeBkg': '{ColorToHex(PrimaryColor)}',
        'nodeTextColor': '{ColorToHex(PrimaryTextColor)}',
        'edgeLabelBackground': '{ColorToHex(BackgroundColor)}',
        'clusterBkg': '{ColorToHex(LightenColor(SecondaryColor, 0.8f))}',
        'clusterBorder': '{ColorToHex(SecondaryColor)}',
        'fillType0': '{ColorToHex(PrimaryColor)}',
        'fillType1': '{ColorToHex(SecondaryColor)}',
        'fillType2': '{ColorToHex(TertiaryColor)}',
        'fillType3': '{ColorToHex(QuaternaryColor)}',
        'fillType4': '{ColorToHex(AccentColor)}',
        'cScale0': '{ColorToHex(PrimaryColor)}',
        'cScale1': '{ColorToHex(SecondaryColor)}',
        'cScale2': '{ColorToHex(TertiaryColor)}'
    }}
}}}}%%";

            return config;
        }

        private string ColorToHex(Color color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        private Color LightenColor(Color color, float factor)
        {
            return Color.FromArgb(
                Math.Min(255, (int)(color.R + (255 - color.R) * factor)),
                Math.Min(255, (int)(color.G + (255 - color.G) * factor)),
                Math.Min(255, (int)(color.B + (255 - color.B) * factor))
            );
        }

        // Create a copy of this palette
        public ColorPalette Clone()
        {
            return new ColorPalette(Name)
            {
                PrimaryColor = PrimaryColor,
                SecondaryColor = SecondaryColor,
                TertiaryColor = TertiaryColor,
                QuaternaryColor = QuaternaryColor,
                PrimaryTextColor = PrimaryTextColor,
                SecondaryTextColor = SecondaryTextColor,
                BackgroundColor = BackgroundColor,
                BorderColor = BorderColor,
                LineColor = LineColor,
                AccentColor = AccentColor
            };
        }
    }
}