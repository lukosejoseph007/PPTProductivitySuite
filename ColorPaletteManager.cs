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

                    // FIXED: Also save as a global "Last Custom" preset that appears in all PowerPoint sessions
                    SaveGlobalCustomPalette(palette);
                }
            }
            catch (Exception ex)
            {
                // Don't show error to user for this non-critical operation
                System.Diagnostics.Debug.WriteLine($"Failed to save last custom palette: {ex.Message}");
            }
        }

        // FIXED: New method to save custom palette globally across all PowerPoint sessions
        private static void SaveGlobalCustomPalette(ColorPalette palette)
        {
            try
            {
                if (palette != null)
                {
                    var globalCustom = palette.Clone();
                    globalCustom.Name = "Last Custom Colors";

                    // Remove any existing "Last Custom Colors" preset
                    _userPresets.RemoveAll(p => p.Name.Equals("Last Custom Colors", StringComparison.OrdinalIgnoreCase));

                    // Add the new custom palette at the beginning of the list
                    _userPresets.Insert(0, globalCustom);

                    // Keep only the most recent 10 presets (including the custom one)
                    if (_userPresets.Count > 10)
                    {
                        _userPresets.RemoveRange(10, _userPresets.Count - 10);
                    }

                    SaveUserPresets();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to save global custom palette: {ex.Message}");
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
                // FIXED: Handle corrupted XML files by backing up and recreating
                System.Diagnostics.Debug.WriteLine($"Failed to load user presets: {ex.Message}");
                
                try
                {
                    // Backup the corrupted file
                    if (File.Exists(PresetsFilePath))
                    {
                        var backupPath = PresetsFilePath + ".backup." + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        File.Copy(PresetsFilePath, backupPath);
                        File.Delete(PresetsFilePath);
                    }
                }
                catch
                {
                    // Ignore backup errors
                }
                
                _userPresets = new List<ColorPalette>();
                
                // Don't show error to user on startup - just log it
                // MessageBox.Show($"User presets were corrupted and have been reset. A backup was created.", "Presets Reset",
                //     MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                // FIXED: Handle corrupted last palette file
                System.Diagnostics.Debug.WriteLine($"Failed to load last custom palette: {ex.Message}");
                
                try
                {
                    // Delete corrupted file
                    if (File.Exists(LastUsedPalettePath))
                    {
                        File.Delete(LastUsedPalettePath);
                    }
                }
                catch
                {
                    // Ignore deletion errors
                }
                
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