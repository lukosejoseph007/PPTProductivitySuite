using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace PPTProductivitySuite
{
    public class ColorPaletteDialog : Form
    {
        public ColorPalette SelectedPalette { get; private set; }
        public bool UseCustomColors { get; private set; }

        private ComboBox cmbPresets;
        private CheckBox chkUseCustomColors;
        private TableLayoutPanel colorPanel;
        private Button[] colorButtons;
        private Label[] colorLabels;
        private Button btnOK;
        private Button btnCancel;
        private Button btnSavePreset;
        private Button btnUpdatePreset;
        private Button btnDeletePreset;
        private ColorDialog colorDialog;

        private ColorPalette currentPalette;
        private bool isCustomPalette = false;
        private readonly string[] colorNames = {
            "Primary", "Secondary", "Tertiary", "Quaternary",
            "Primary Text", "Secondary Text", "Background",
            "Border", "Line", "Accent"
        };

        public ColorPaletteDialog(ColorPalette initialPalette = null)
        {
            // FIXED: Properly preserve the initial palette instead of always defaulting
            if (initialPalette != null)
            {
                currentPalette = initialPalette.Clone();
                isCustomPalette = !IsBuiltInPreset(initialPalette.Name);
            }
            else
            {
                // Try to load last custom palette first
                var lastCustom = ColorPaletteManager.GetLastCustomPalette();
                if (lastCustom != null)
                {
                    currentPalette = lastCustom;
                    isCustomPalette = true;
                }
                else
                {
                    currentPalette = ColorPaletteManager.GetPreset("Corporate Blue");
                    isCustomPalette = false;
                }
            }

            InitializeComponents();
            LoadPresets();
            UpdateColorDisplay();
        }

        private bool IsBuiltInPreset(string presetName)
        {
            return ColorPalette.GetBuiltInPresets().ContainsKey(presetName);
        }

        private void InitializeComponents()
        {
            Text = "Select Color Palette";
            Size = new Size(900, 750); // FIXED: Much larger size for better layout
            MinimumSize = new Size(850, 700); // FIXED: Larger minimum size
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            ShowIcon = false;

            colorDialog = new ColorDialog
            {
                FullOpen = true,
                AllowFullOpen = true
            };

            // Main layout panel
            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(20) // FIXED: More padding around main panel
            };
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 150)); // FIXED: More space for text wrapping
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 100)); // FIXED: More space for taller save buttons
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 90)); // FIXED: More space for bottom buttons

            // Preset selection panel - FIXED button sizing
            var presetPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10) // FIXED: More padding in preset panel
            };

            var lblPreset = new Label
            {
                Text = "Color Preset:",
                Location = new Point(10, 15), // FIXED: More margin from edge
                Size = new Size(120, 30),
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F)
            };

            cmbPresets = new ComboBox
            {
                Location = new Point(140, 12), // FIXED: More space from label
                Size = new Size(280, 32), // FIXED: Even wider
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F)
            };

            // FIXED: Better button positioning with more spacing
            btnDeletePreset = new Button
            {
                Text = "Delete Preset",
                Location = new Point(430, 10), // FIXED: More space from combobox
                Size = new Size(200, 35),
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false
            };

            chkUseCustomColors = new CheckBox
            {
                Text = "Use custom colors (allows editing)",
                Location = new Point(10, 60),
                Size = new Size(400, 40), // FIXED: Much wider and taller for text wrapping
                Checked = isCustomPalette,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false, // Keep fixed size but allow text wrapping
                UseCompatibleTextRendering = true
            };

            presetPanel.Controls.AddRange(new Control[] { lblPreset, cmbPresets, btnDeletePreset, chkUseCustomColors });

            // Color selection panel - improved layout
            colorPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 5,
                RowCount = 4,
                Padding = new Padding(15), // FIXED: More padding around color grid
                Enabled = chkUseCustomColors.Checked
            };

            // Set column and row styles for better distribution
            for (int i = 0; i < 5; i++)
            {
                colorPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            }
            for (int i = 0; i < 4; i++)
            {
                colorPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 25));
            }

            colorButtons = new Button[10];
            colorLabels = new Label[10];

            for (int i = 0; i < 10; i++)
            {
                colorLabels[i] = new Label
                {
                    Text = colorNames[i],
                    TextAlign = ContentAlignment.MiddleCenter,
                    Dock = DockStyle.Fill,
                    Font = new Font(SystemFonts.DefaultFont.FontFamily, 8F),
                    Margin = new Padding(5) // FIXED: More margin around labels
                };

                colorButtons[i] = new Button
                {
                    Dock = DockStyle.Fill,
                    FlatStyle = FlatStyle.Flat,
                    MinimumSize = new Size(110, 65),
                    Tag = i,
                    Margin = new Padding(5) // FIXED: More margin around color buttons
                };
                colorButtons[i].Click += ColorButton_Click;

                int row = (i / 5) * 2; // Each color takes 2 rows (label + button)
                int col = i % 5;

                colorPanel.Controls.Add(colorLabels[i], col, row);
                colorPanel.Controls.Add(colorButtons[i], col, row + 1);
            }

            // Button panel for save preset - FIXED button sizing
            var savePanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15) // FIXED: More padding in save panel
            };
            btnSavePreset = new Button
            {
                Text = "Save as New Preset...",
                Location = new Point(15, 20),
                Size = new Size(200, 50), // FIXED: Taller to prevent text wrapping
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                Enabled = chkUseCustomColors.Checked,
                AutoSize = false,
                UseCompatibleTextRendering = true
            };

            // FIXED: Add Update Preset button with proper spacing
            var btnUpdatePreset = new Button
            {
                Text = "Update Selected Preset",
                Location = new Point(230, 20),
                Size = new Size(200, 50), // FIXED: Taller to prevent text wrapping
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                Enabled = false,
                AutoSize = false,
                UseCompatibleTextRendering = true
            };
            btnUpdatePreset.Click += BtnUpdatePreset_Click;

            savePanel.Controls.AddRange(new Control[] { btnSavePreset, btnUpdatePreset });

            // Store reference for enabling/disabling
            this.btnUpdatePreset = btnUpdatePreset;

            // Bottom button panel - FIXED button sizing
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(15, 25, 15, 15) // FIXED: More padding around buttons
            };

            btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Size = new Size(120, 45), // FIXED: Larger buttons
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F)
            };

            btnOK = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Size = new Size(120, 45), // FIXED: Larger buttons
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F)
            };

            buttonPanel.Controls.AddRange(new Control[] { btnCancel, btnOK });

            // Add to main panel
            mainPanel.Controls.Add(presetPanel, 0, 0);
            mainPanel.Controls.Add(colorPanel, 0, 1);
            mainPanel.Controls.Add(savePanel, 0, 2);
            mainPanel.Controls.Add(buttonPanel, 0, 3);

            Controls.Add(mainPanel);

            AcceptButton = btnOK;
            CancelButton = btnCancel;

            // Event handlers
            cmbPresets.SelectedIndexChanged += CmbPresets_SelectedIndexChanged;
            chkUseCustomColors.CheckedChanged += ChkUseCustomColors_CheckedChanged;
            btnOK.Click += BtnOK_Click;
            btnSavePreset.Click += BtnSavePreset_Click;
            btnDeletePreset.Click += BtnDeletePreset_Click;
        }

        private void LoadPresets()
        {
            cmbPresets.Items.Clear();
            var allPresets = ColorPaletteManager.GetAllPresets();

            foreach (var preset in allPresets)
            {
                cmbPresets.Items.Add(preset.Name);
            }

            // FIXED: Properly select the current palette or mark as custom
            if (cmbPresets.Items.Count > 0)
            {
                var currentName = currentPalette?.Name;
                if (!string.IsNullOrEmpty(currentName) && cmbPresets.Items.Contains(currentName))
                {
                    cmbPresets.SelectedItem = currentName;
                    isCustomPalette = false;
                }
                else
                {
                    // If current palette doesn't match any preset, it's custom
                    // But don't clear the selection immediately - let user choose
                    if (cmbPresets.Items.Count > 0)
                    {
                        cmbPresets.SelectedIndex = 0; // Select first preset as default
                    }
                    isCustomPalette = true;
                    chkUseCustomColors.Checked = true;
                }
            }

            UpdateDeleteButtonState();
            UpdateUpdateButtonState();
        }

        private void UpdateColorDisplay()
        {
            if (currentPalette == null) return;

            var colors = new Color[]
            {
                currentPalette.PrimaryColor,
                currentPalette.SecondaryColor,
                currentPalette.TertiaryColor,
                currentPalette.QuaternaryColor,
                currentPalette.PrimaryTextColor,
                currentPalette.SecondaryTextColor,
                currentPalette.BackgroundColor,
                currentPalette.BorderColor,
                currentPalette.LineColor,
                currentPalette.AccentColor
            };

            for (int i = 0; i < colorButtons.Length; i++)
            {
                colorButtons[i].BackColor = colors[i];
                colorButtons[i].ForeColor = GetContrastColor(colors[i]);
            }
        }

        private Color GetContrastColor(Color backgroundColor)
        {
            var brightness = (backgroundColor.R * 299 + backgroundColor.G * 587 + backgroundColor.B * 114) / 1000;
            return brightness > 128 ? Color.Black : Color.White;
        }

        private void ColorButton_Click(object sender, EventArgs e)
        {
            if (!chkUseCustomColors.Checked) return;

            var button = sender as Button;
            var colorIndex = (int)button.Tag;

            colorDialog.Color = button.BackColor;
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                button.BackColor = colorDialog.Color;
                button.ForeColor = GetContrastColor(colorDialog.Color);
                UpdateCurrentPaletteFromButtons();

                // FIXED: Mark as custom when colors are modified but don't clear selection immediately
                // This allows users to update existing presets
                if (cmbPresets.SelectedIndex == -1)
                {
                    isCustomPalette = true;
                }

                // FIXED: Automatically save custom colors globally for persistence across PowerPoint sessions
                if (isCustomPalette || cmbPresets.SelectedIndex == -1)
                {
                    ColorPaletteManager.SaveLastCustomPalette(currentPalette);
                }

                UpdateDeleteButtonState();
                UpdateUpdateButtonState();
            }
        }

        private void UpdateCurrentPaletteFromButtons()
        {
            currentPalette.PrimaryColor = colorButtons[0].BackColor;
            currentPalette.SecondaryColor = colorButtons[1].BackColor;
            currentPalette.TertiaryColor = colorButtons[2].BackColor;
            currentPalette.QuaternaryColor = colorButtons[3].BackColor;
            currentPalette.PrimaryTextColor = colorButtons[4].BackColor;
            currentPalette.SecondaryTextColor = colorButtons[5].BackColor;
            currentPalette.BackgroundColor = colorButtons[6].BackColor;
            currentPalette.BorderColor = colorButtons[7].BackColor;
            currentPalette.LineColor = colorButtons[8].BackColor;
            currentPalette.AccentColor = colorButtons[9].BackColor;

            // FIXED: Update the name to indicate it's custom
            if (isCustomPalette)
            {
                currentPalette.Name = "Custom";
            }
        }

        private void CmbPresets_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbPresets.SelectedItem != null)
            {
                var presetName = cmbPresets.SelectedItem.ToString();
                var selectedPreset = ColorPaletteManager.GetPreset(presetName);
                if (selectedPreset != null)
                {
                    currentPalette = selectedPreset;
                    isCustomPalette = false;
                    UpdateColorDisplay();
                    UpdateDeleteButtonState();
                    UpdateUpdateButtonState();
                }
            }
        }

        private void ChkUseCustomColors_CheckedChanged(object sender, EventArgs e)
        {
            colorPanel.Enabled = chkUseCustomColors.Checked;
            btnSavePreset.Enabled = chkUseCustomColors.Checked;
            UpdateUpdateButtonState(); // FIXED: Update the Update button state too

            // FIXED: Don't clear preset selection when enabling custom colors
            // Just mark as custom if no preset is selected
            if (chkUseCustomColors.Checked && cmbPresets.SelectedIndex == -1)
            {
                isCustomPalette = true;
            }
            // FIXED: When disabling custom colors, ensure we're not in custom mode
            else if (!chkUseCustomColors.Checked && isCustomPalette)
            {
                // If we were in custom mode, select the first preset
                if (cmbPresets.Items.Count > 0)
                {
                    cmbPresets.SelectedIndex = 0;
                }
            }
        }

        private void UpdateDeleteButtonState()
        {
            var builtInPresets = ColorPalette.GetBuiltInPresets().Keys;
            var selectedPreset = cmbPresets.SelectedItem?.ToString();
            btnDeletePreset.Enabled = selectedPreset != null && !builtInPresets.Contains(selectedPreset);
        }

        // FIXED: New method to control Update button state
        private void UpdateUpdateButtonState()
        {
            var builtInPresets = ColorPalette.GetBuiltInPresets().Keys;
            var selectedPreset = cmbPresets.SelectedItem?.ToString();
            
            // Enable update button if:
            // 1. Custom colors is checked
            // 2. A preset is selected
            // 3. The selected preset is not a built-in preset
            btnUpdatePreset.Enabled = chkUseCustomColors.Checked &&
                                     !string.IsNullOrEmpty(selectedPreset) &&
                                     !builtInPresets.Contains(selectedPreset);
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            // FIXED: Ensure we return the current state properly
            SelectedPalette = currentPalette?.Clone();
            UseCustomColors = chkUseCustomColors.Checked;

            // Ensure the returned palette has the correct name
            if (SelectedPalette != null && isCustomPalette)
            {
                SelectedPalette.Name = "Custom";
                // Save custom palette for future use
                ColorPaletteManager.SaveLastCustomPalette(SelectedPalette);
            }
        }

        private void BtnSavePreset_Click(object sender, EventArgs e)
        {
            using (var nameDialog = new PresetNameDialog())
            {
                if (nameDialog.ShowDialog() == DialogResult.OK)
                {
                    var presetName = nameDialog.PresetName;

                    // Check if name already exists and offer to create unique name
                    if (ColorPaletteManager.PresetNameExists(presetName))
                    {
                        var uniqueName = ColorPaletteManager.GetUniquePresetName(presetName);
                        var result = MessageBox.Show($"A preset named '{presetName}' already exists. Save as '{uniqueName}' instead?",
                            "Name Exists", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                        {
                            presetName = uniqueName;
                        }
                        else if (result == DialogResult.No)
                        {
                            // User wants to overwrite
                            var overwriteResult = MessageBox.Show($"Do you want to overwrite the existing preset '{presetName}'?",
                                "Confirm Overwrite", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                            if (overwriteResult != DialogResult.Yes)
                                return;
                        }
                        else
                        {
                            return; // Cancel
                        }
                    }

                    var newPreset = currentPalette.Clone();
                    newPreset.Name = presetName;

                    if (ColorPaletteManager.SaveUserPreset(newPreset))
                    {
                        LoadPresets();
                        cmbPresets.SelectedItem = newPreset.Name;
                        isCustomPalette = false; // It's now a saved preset
                        MessageBox.Show($"Preset '{newPreset.Name}' saved successfully!", "Success",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        // FIXED: New method to update existing presets
        private void BtnUpdatePreset_Click(object sender, EventArgs e)
        {
            var selectedPreset = cmbPresets.SelectedItem?.ToString();
            if (selectedPreset == null)
            {
                MessageBox.Show("Please select a preset to update.", "No Preset Selected",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var builtInPresets = ColorPalette.GetBuiltInPresets().Keys;
            if (builtInPresets.Contains(selectedPreset))
            {
                MessageBox.Show("Built-in presets cannot be updated.", "Cannot Update",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var result = MessageBox.Show($"Are you sure you want to update the preset '{selectedPreset}' with the current colors?",
                "Confirm Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                var updatedPreset = currentPalette.Clone();
                updatedPreset.Name = selectedPreset;

                if (ColorPaletteManager.SaveUserPreset(updatedPreset))
                {
                    MessageBox.Show($"Preset '{selectedPreset}' updated successfully!", "Success",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    // Refresh the display to show updated colors
                    LoadPresets();
                    cmbPresets.SelectedItem = selectedPreset;
                }
            }
        }



        private void BtnDeletePreset_Click(object sender, EventArgs e)
        {
            var selectedPreset = cmbPresets.SelectedItem?.ToString();
            if (selectedPreset == null) return;

            var result = MessageBox.Show($"Are you sure you want to delete the preset '{selectedPreset}'?",
                "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                if (ColorPaletteManager.DeleteUserPreset(selectedPreset))
                {
                    LoadPresets();
                    MessageBox.Show($"Preset '{selectedPreset}' deleted successfully!", "Success",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
    }

    // FIXED: Improved PresetNameDialog with better button sizing
    public class PresetNameDialog : Form
    {
        public string PresetName { get; private set; }

        private TextBox txtName;
        private Button btnOK;
        private Button btnCancel;

        public PresetNameDialog()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            Text = "Save Preset";
            Size = new Size(420, 250); // FIXED: Larger for better layout and spacing
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;

            var lblName = new Label
            {
                Text = "Preset Name:",
                Location = new Point(20, 30), // FIXED: More margin from edge
                Size = new Size(120, 20),
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F)
            };

            txtName = new TextBox
            {
                Location = new Point(20, 55), // FIXED: More vertical spacing
                Size = new Size(370, 28), // FIXED: Wider text box with more margin
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F)
            };

            // FIXED: Better button positioning with more spacing
            btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(310, 110), // FIXED: More vertical spacing
                Size = new Size(80, 35),
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F)
            };

            btnOK = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new Point(220, 110), // FIXED: More spacing between buttons
                Size = new Size(80, 35),
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F)
            };

            Controls.AddRange(new Control[] { lblName, txtName, btnCancel, btnOK });

            AcceptButton = btnOK;
            CancelButton = btnCancel;

            btnOK.Click += (s, e) =>
            {
                var enteredName = txtName.Text.Trim();
                if (string.IsNullOrWhiteSpace(enteredName))
                {
                    MessageBox.Show("Please enter a preset name.", "Name Required",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtName.Focus();
                    return;
                }

                // FIXED: Add basic name validation
                if (enteredName.Length > 50)
                {
                    MessageBox.Show("Preset name is too long (maximum 50 characters).", "Name Too Long",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtName.Focus();
                    return;
                }

                // Check for invalid characters
                if (enteredName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                {
                    MessageBox.Show("Preset name contains invalid characters.", "Invalid Name",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtName.Focus();
                    return;
                }

                PresetName = enteredName;
            };

            txtName.Focus();
        }
    }
}