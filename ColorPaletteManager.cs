using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;
using System.Windows.Forms;
using System.Linq;

namespace PPTProductivitySuite
{
    public static class ColorPaletteManager
    {
        private static readonly string PresetsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PPTProductivitySuite",
            "ColorPresets.xml"
        );

        // FIXED: Increase limit for user presets and track last used custom palette
        private static readonly string LastUsedPalettePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PPTProductivitySuite",
            "LastUsedPalette.xml"
        );

        private static List<ColorPalette> _userPresets;
        private static ColorPalette _lastCustomPalette;

        static ColorPaletteManager()
        {
            EnsureDirectoryExists();
            LoadUserPresets();
            LoadLastCustomPalette();
        }

        public static List<ColorPalette> GetAllPresets()
        {
            var allPresets = new List<ColorPalette>();

            // Add built-in presets
            foreach (var preset in ColorPalette.GetBuiltInPresets().Values)
            {
                allPresets.Add(preset);
            }

            // Add user presets
            allPresets.AddRange(_userPresets);

            return allPresets;
        }

        public static List<ColorPalette> GetUserPresets()
        {
            return new List<ColorPalette>(_userPresets);
        }

        // FIXED: New method to save/retrieve last used custom palette
        public static ColorPalette GetLastCustomPalette()
        {
            return _lastCustomPalette?.Clone();
        }

        public static void SaveLastCustomPalette(ColorPalette palette)
        {
            try
            {
                if (palette != null)
                {
                    _lastCustomPalette = palette.Clone();
                    _lastCustomPalette.Name = "Custom"; // Ensure it has the right name

                    var serializer = new XmlSerializer(typeof(ColorPalette));
                    using (var writer = new StreamWriter(LastUsedPalettePath))
                    {
                        serializer.Serialize(writer, _lastCustomPalette);
                    }
                }
            }
            catch (Exception ex)
            {
                // Don't show error to user for this non-critical operation
                System.Diagnostics.Debug.WriteLine($"Failed to save last custom palette: {ex.Message}");
            }
        }

        public static bool SaveUserPreset(ColorPalette palette)
        {
            try
            {
                // Remove existing preset with same name (case-insensitive)
                _userPresets.RemoveAll(p => p.Name.Equals(palette.Name, StringComparison.OrdinalIgnoreCase));

                // Add new preset
                _userPresets.Add(palette.Clone());

                // FIXED: Increase limit to 10 user presets and use LRU (Least Recently Used) strategy
                if (_userPresets.Count > 10)
                {
                    // Remove the oldest preset (first in list)
                    _userPresets.RemoveAt(0);
                }

                SaveUserPresets();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save preset: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public static bool DeleteUserPreset(string name)
        {
            try
            {
                var removed = _userPresets.RemoveAll(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                if (removed > 0)
                {
                    SaveUserPresets();
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to delete preset: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public static ColorPalette GetPreset(string name)
        {
            // Handle special case for "Custom" - return last custom palette if available
            if (name == "Custom" && _lastCustomPalette != null)
            {
                return _lastCustomPalette.Clone();
            }

            // Check built-in presets first
            var builtIn = ColorPalette.GetBuiltInPresets();
            if (builtIn.ContainsKey(name))
            {
                return builtIn[name].Clone();
            }

            // Check user presets
            var userPreset = _userPresets.Find(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (userPreset != null)
            {
                // FIXED: Move accessed preset to end (LRU strategy)
                _userPresets.Remove(userPreset);
                _userPresets.Add(userPreset);
                SaveUserPresets(); // Save the reordering
                return userPreset.Clone();
            }

            // Return null if not found
            return null;
        }

        // FIXED: New method to check if a preset name already exists
        public static bool PresetNameExists(string name)
        {
            var builtIn = ColorPalette.GetBuiltInPresets();
            if (builtIn.ContainsKey(name))
                return true;

            return _userPresets.Any(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        // FIXED: New method to get a unique preset name
        public static string GetUniquePresetName(string baseName)
        {
            if (!PresetNameExists(baseName))
                return baseName;

            int counter = 1;
            string uniqueName;
            do
            {
                uniqueName = $"{baseName} ({counter})";
                counter++;
            } while (PresetNameExists(uniqueName));

            return uniqueName;
        }

        private static void EnsureDirectoryExists()
        {
            var directory = Path.GetDirectoryName(PresetsFilePath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
        }

        private static void LoadUserPresets()
        {
            _userPresets = new List<ColorPalette>();

            try
            {
                if (File.Exists(PresetsFilePath))
                {
                    var serializer = new XmlSerializer(typeof(List<ColorPalette>));
                    using (var reader = new StreamReader(PresetsFilePath))
                    {
                        _userPresets = (List<ColorPalette>)serializer.Deserialize(reader) ?? new List<ColorPalette>();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load user presets: {ex.Message}", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _userPresets = new List<ColorPalette>();
            }
        }

        private static void LoadLastCustomPalette()
        {
            try
            {
                if (File.Exists(LastUsedPalettePath))
                {
                    var serializer = new XmlSerializer(typeof(ColorPalette));
                    using (var reader = new StreamReader(LastUsedPalettePath))
                    {
                        _lastCustomPalette = (ColorPalette)serializer.Deserialize(reader);
                    }
                }
            }
            catch (Exception ex)
            {
                // Non-critical error, just log it
                System.Diagnostics.Debug.WriteLine($"Failed to load last custom palette: {ex.Message}");
                _lastCustomPalette = null;
            }
        }

        private static void SaveUserPresets()
        {
            try
            {
                var serializer = new XmlSerializer(typeof(List<ColorPalette>));
                using (var writer = new StreamWriter(PresetsFilePath))
                {
                    serializer.Serialize(writer, _userPresets);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to save presets: {ex.Message}", ex);
            }
        }
    }
}