using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace PPTProductivitySuite
{
    public partial class MermaidInputDialog : Form
    {
        public string MermaidCode { get; private set; }
        public bool UseCustomColors { get; private set; } = true;
        public ColorPalette SelectedColorPalette { get; private set; }

        private TextBox txtMermaidCode;
        private Label lblInstructions;
        private Button btnInsert;
        private Button btnCancel;
        private Button btnExample;
        private CheckBox chkUseColors;
        private Button btnSelectColors;
        private Label lblCurrentPalette;
        private Panel pnlColorControls;

        // FIXED: Static field to preserve palette across dialog instances
        private static ColorPalette _lastUsedPalette = null;

        public MermaidInputDialog()
        {
            // FIXED: Initialize with last used palette or default
            if (_lastUsedPalette != null)
            {
                SelectedColorPalette = _lastUsedPalette.Clone();
            }
            else
            {
                SelectedColorPalette = ColorPaletteManager.GetPreset("Corporate Blue");
            }

            InitializeCustomComponents();
            this.Text = "Insert Mermaid Diagram";
            this.StartPosition = FormStartPosition.CenterParent;
            this.MinimumSize = new Size(850, 700); // Larger for better layout
            this.Size = new Size(950, 750); // Larger default size
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
        }

        private void InitializeCustomComponents()
        {
            // Instructions label - FIXED: Better sizing
            lblInstructions = new Label
            {
                Text = "Enter your Mermaid diagram code below. You can customize colors using the color palette selector.",
                Dock = DockStyle.Top,
                Height = 70, // Increased height for better text display
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(15, 15, 15, 5),
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false // Prevent auto-sizing issues
            };

            // Color controls panel - FIXED: Better layout
            pnlColorControls = new Panel
            {
                Dock = DockStyle.Top,
                Height = 100, // Increased height
                Padding = new Padding(15, 10, 15, 10)
            };

            // Use colors checkbox - FIXED: Better positioning and sizing
            chkUseColors = new CheckBox
            {
                Text = "Use Custom Color Palette",
                Location = new Point(15, 15),
                Size = new Size(220, 25), // Wider for full text
                Checked = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F)
            };

            // Select colors button - FIXED: Better sizing and positioning
            btnSelectColors = new Button
            {
                Text = "Select Colors...",
                Location = new Point(250, 12),
                Size = new Size(130, 35), // Larger button
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false // Prevent auto-sizing issues
            };

            // Current palette label - FIXED: Better positioning and text handling
            lblCurrentPalette = new Label
            {
                Text = $"Current Palette: {SelectedColorPalette?.Name ?? "Default"}",
                Location = new Point(15, 55),
                Size = new Size(500, 30), // Wider and taller for longer text
                ForeColor = Color.FromArgb(100, 100, 100),
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 8.5F),
                AutoSize = false, // Prevent auto-sizing issues
                AutoEllipsis = true // Show ellipsis for long text
            };

            // Add controls to color panel
            pnlColorControls.Controls.AddRange(new Control[] { chkUseColors, btnSelectColors, lblCurrentPalette });

            // Main text box with proper container
            var textContainer = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(15, 10, 15, 10)
            };

            txtMermaidCode = new TextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Vertical,
                AcceptsTab = true,
                Font = new Font("Consolas", 11F, FontStyle.Regular),
                Text = GetDefaultExample(),
                WordWrap = false
            };

            textContainer.Controls.Add(txtMermaidCode);

            // Buttons - FIXED: Better sizing to prevent text overflow
            btnInsert = new Button
            {
                Text = "Insert Diagram",
                DialogResult = DialogResult.OK,
                Size = new Size(140, 40), // Larger for better text fit
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false // Prevent auto-sizing issues
            };

            btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Size = new Size(100, 40), // Larger buttons
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false
            };

            btnExample = new Button
            {
                Text = "Examples...",
                Size = new Size(120, 40), // Larger for better text fit
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false
            };

            // Button panel - FIXED: Better spacing
            var flowPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 70, // Increased height
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(15, 15, 15, 15) // More padding
            };

            flowPanel.Controls.Add(btnCancel);
            flowPanel.Controls.Add(btnInsert);
            flowPanel.Controls.Add(btnExample);

            // Form setup
            this.Controls.Add(textContainer);
            this.Controls.Add(flowPanel);
            this.Controls.Add(pnlColorControls);
            this.Controls.Add(lblInstructions);
            this.AcceptButton = btnInsert;
            this.CancelButton = btnCancel;
            this.ShowIcon = false;

            // Event handlers
            btnInsert.Click += (s, e) =>
            {
                MermaidCode = txtMermaidCode.Text.Trim();
                if (string.IsNullOrWhiteSpace(MermaidCode))
                {
                    MessageBox.Show("Please enter Mermaid code before inserting.", "No Code Entered",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                UseCustomColors = chkUseColors.Checked;

                // FIXED: Save the palette for next time
                if (SelectedColorPalette != null)
                {
                    _lastUsedPalette = SelectedColorPalette.Clone();
                }
            };

            btnExample.Click += BtnExample_Click;
            btnSelectColors.Click += BtnSelectColors_Click;

            chkUseColors.CheckedChanged += (s, e) =>
            {
                btnSelectColors.Enabled = chkUseColors.Checked;
                lblCurrentPalette.Enabled = chkUseColors.Checked;
            };

            // Select all text when dialog opens
            this.Shown += (s, e) => txtMermaidCode.SelectAll();
        }

        private void BtnSelectColors_Click(object sender, EventArgs e)
        {
            // FIXED: Pass the current palette to preserve custom colors
            using (var colorDialog = new ColorPaletteDialog(SelectedColorPalette))
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    SelectedColorPalette = colorDialog.SelectedPalette;
                    // FIXED: Update label with proper text handling
                    var paletteName = SelectedColorPalette?.Name ?? "Custom";
                    lblCurrentPalette.Text = $"Current Palette: {paletteName}";

                    // FIXED: Save the palette for future use
                    _lastUsedPalette = SelectedColorPalette?.Clone();
                }
            }
        }

        private void BtnExample_Click(object sender, EventArgs e)
        {
            var examples = new[]
            {
               GetDefaultExample(),
               GetFlowchartExample(),
               GetSequenceDiagramExample(),
               GetGanttExample(),
               GetClassDiagramExample()
           };

            using (var exampleForm = new ExampleSelectionForm(examples))
            {
                if (exampleForm.ShowDialog() == DialogResult.OK)
                {
                    txtMermaidCode.Text = exampleForm.SelectedExample;
                    txtMermaidCode.SelectAll();
                    txtMermaidCode.Focus();
                }
            }
        }

        private string GetDefaultExample()
        {
            return @"graph TD
   A[Start] --> B{Decision}
   B -->|Yes| C[Process 1]
   B -->|No| D[Process 2]
   C --> E[End]
   D --> E";
        }

        private string GetFlowchartExample()
        {
            return @"flowchart LR
   A[Square Rect] --> B((Circle))
   A --> C(Round Rect)
   B --> D{Rhombus}
   C --> D";
        }

        private string GetSequenceDiagramExample()
        {
            return @"sequenceDiagram
   participant Alice
   participant Bob
   Alice->>John: Hello John, how are you?
   loop Healthcheck
       John->>John: Fight against hypochondria
   end
   Note right of John: Rational thoughts <br/>prevail!
   John-->>Alice: Great!
   John->>Bob: How about you?
   Bob-->>John: Jolly good!";
        }

        private string GetGanttExample()
        {
            return @"gantt
   title A Gantt Diagram
   dateFormat  YYYY-MM-DD
   section Section
   A task           :a1, 2014-01-01, 30d
   Another task     :after a1  , 20d
   section Another
   Task in sec      :2014-01-12  , 12d
   another task      : 24d";
        }

        private string GetClassDiagramExample()
        {
            return @"classDiagram
   Class01 <|-- AveryLongClass : Cool
   Class03 *-- Class04
   Class05 o-- Class06
   Class07 .. Class08
   Class09 --> C2 : Where am i?
   Class09 --* C3
   Class09 --|> Class07
   Class07 : equals()
   Class07 : Object[] elementData
   Class01 : size()
   Class01 : int chimp
   Class01 : int gorilla";
        }
    }

    // FIXED: Improved ExampleSelectionForm with better sizing and layout
    public class ExampleSelectionForm : Form
    {
        public string SelectedExample { get; private set; }
        private readonly string[] _examples;

        public ExampleSelectionForm(string[] examples)
        {
            _examples = examples;
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            Text = "Select Example";
            Size = new Size(550, 450); // Increased size for better layout
            StartPosition = FormStartPosition.CenterParent;
            ShowIcon = false;
            MaximizeBox = false;
            MinimizeBox = false;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            var listBox = new ListBox
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(15),
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 10F),
                ItemHeight = 25 // Taller items for better readability
            };

            listBox.Items.Add("Basic Graph");
            listBox.Items.Add("Flowchart");
            listBox.Items.Add("Sequence Diagram");
            listBox.Items.Add("Gantt Chart");
            listBox.Items.Add("Class Diagram");
            listBox.SelectedIndex = 0;

            // FIXED: Better button sizing
            var btnOK = new Button
            {
                Text = "Select",
                DialogResult = DialogResult.OK,
                Size = new Size(120, 40), // Larger buttons
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false
            };
            var btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Size = new Size(100, 40), // Larger buttons
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false
            };

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 70, // Increased height
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(15, 15, 15, 15) // More padding
            };

            buttonPanel.Controls.Add(btnCancel);
            buttonPanel.Controls.Add(btnOK);

            Controls.Add(listBox);
            Controls.Add(buttonPanel);

            AcceptButton = btnOK;
            CancelButton = btnCancel;

            btnOK.Click += (s, e) =>
            {
                if (listBox.SelectedIndex >= 0)
                    SelectedExample = _examples[listBox.SelectedIndex];
            };
        }
    }
}