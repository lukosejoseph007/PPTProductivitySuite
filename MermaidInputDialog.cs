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
            this.MinimumSize = new Size(1000, 800); // FIXED: Much larger for better layout
            this.Size = new Size(1200, 900); // FIXED: Much larger default size
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
                Height = 80, // FIXED: Even more height for better text display
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(20, 20, 20, 10), // FIXED: More padding around instructions
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false // Prevent auto-sizing issues
            };

            // Color controls panel - FIXED: Better layout
            pnlColorControls = new Panel
            {
                Dock = DockStyle.Top,
                Height = 140, // FIXED: Even more height for color controls
                Padding = new Padding(20, 15, 20, 15) // FIXED: More padding around color controls
            };

            // Use colors checkbox - FIXED: Better positioning and sizing
            chkUseColors = new CheckBox
            {
                Text = "Use Custom Color Palette",
                Location = new Point(20, 20),
                Size = new Size(250, 35), // FIXED: Taller to prevent text wrapping
                Checked = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false,
                UseCompatibleTextRendering = true
            };

            // Select colors button - FIXED: Better sizing and positioning
            btnSelectColors = new Button
            {
                Text = "Select Colors...",
                Location = new Point(280, 17), // FIXED: More space from checkbox
                Size = new Size(150, 40), // FIXED: Larger button
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false // Prevent auto-sizing issues
            };

            // Current palette label - FIXED: Better positioning and text handling
            lblCurrentPalette = new Label
            {
                Text = $"Current Palette: {SelectedColorPalette?.Name ?? "Default"}",
                Location = new Point(20, 70), // FIXED: More vertical spacing
                Size = new Size(600, 40), // FIXED: Even taller for longer text to prevent wrapping
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
                Padding = new Padding(20, 15, 20, 15) // FIXED: More padding around text container
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
                Size = new Size(160, 50), // FIXED: Taller to prevent text wrapping
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false // Prevent auto-sizing issues
            };

            btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Size = new Size(120, 50), // FIXED: Taller buttons to prevent text wrapping
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false
            };

            btnExample = new Button
            {
                Text = "Examples...",
                Size = new Size(140, 50), // FIXED: Taller to prevent text wrapping
                UseVisualStyleBackColor = true,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false
            };

            // Button panel - FIXED: Better spacing
            var flowPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 90, // FIXED: Even more height for taller buttons
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(20, 20, 20, 20) // FIXED: Even more padding around buttons
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
               GetClassDiagramExample(),
               GetERDiagramExample(),
               GetMindMapExample(),
               GetStateDiagramExample(),
               GetUserJourneyExample()
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
    participant John
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
    title Project Timeline
    dateFormat YYYY-MM-DD
    section Planning
    Requirements    :done, req, 2024-01-01, 2024-01-15
    Design          :done, design, after req, 10d
    section Development
    Backend         :active, backend, 2024-01-25, 20d
    Frontend        :frontend, after backend, 15d
    section Testing
    Unit Tests      :testing, after frontend, 10d
    Integration     :integration, after testing, 5d";
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

        private string GetERDiagramExample()
        {
            return @"erDiagram
    CUSTOMER {
        string customer_id PK
        string first_name
        string last_name
        string email
        date created_at
    }
    ORDER {
        string order_id PK
        string customer_id FK
        decimal total_amount
        date order_date
        string status
    }
    PRODUCT {
        string product_id PK
        string name
        decimal price
        int stock_quantity
        string category
    }
    ORDER_ITEM {
        string order_id FK
        string product_id FK
        int quantity
        decimal unit_price
    }
    
    CUSTOMER ||--o{ ORDER : places
    ORDER ||--o{ ORDER_ITEM : contains
    PRODUCT ||--o{ ORDER_ITEM : ""ordered in""";
        }

        private string GetMindMapExample()
        {
            return @"mindmap
  root((Project Planning))
    Requirements
      Functional
        User Stories
        Use Cases
      Non-Functional
        Performance
        Security
        Scalability
    Design
      Architecture
        Frontend
        Backend
        Database
      UI/UX
        Wireframes
        Mockups
        User Flow
    Development
      Frontend
        React
        CSS
      Backend
        API
        Database
      Testing
        Unit Tests
        Integration Tests
    Deployment
      CI/CD
      Production
      Monitoring";
        }

        private string GetStateDiagramExample()
        {
            return @"stateDiagram-v2
    [*] --> Idle
    Idle --> Processing : start_process
    Processing --> Success : process_complete
    Processing --> Error : process_failed
    Success --> [*]
    Error --> Retry : retry_process
    Error --> [*] : give_up
    Retry --> Processing : attempt_again
    Retry --> [*] : max_retries_reached
    
    state Processing {
        [*] --> Validating
        Validating --> Executing : validation_passed
        Validating --> [*] : validation_failed
        Executing --> [*] : execution_complete
    }";
        }

        private string GetUserJourneyExample()
        {
            return @"journey
    title User Shopping Journey
    section Discovery
      Visit website     : 5: User
      Browse products   : 4: User
      Read reviews      : 3: User
    section Selection
      Compare items     : 4: User
      Add to cart       : 5: User
      Check inventory   : 3: User, System
    section Purchase
      Enter details     : 2: User
      Process payment   : 1: User, System
      Confirm order     : 5: User, System
    section Fulfillment
      Pack order        : 3: Staff
      Ship product      : 4: Staff, Courier
      Deliver package   : 5: Courier
      Receive product   : 5: User";
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
            Size = new Size(650, 550); // FIXED: Much larger size for better layout
            StartPosition = FormStartPosition.CenterParent;
            ShowIcon = false;
            MaximizeBox = false;
            MinimizeBox = false;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            var listBox = new ListBox
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(20), // FIXED: More margin around listbox
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 10F),
                ItemHeight = 30
            };

            listBox.Items.Add("Basic Graph");
            listBox.Items.Add("Flowchart");
            listBox.Items.Add("Sequence Diagram");
            listBox.Items.Add("Gantt Chart");
            listBox.Items.Add("Class Diagram");
            listBox.Items.Add("Entity Relationship Diagram");
            listBox.Items.Add("Mind Map");
            listBox.Items.Add("State Diagram");
            listBox.Items.Add("User Journey");
            listBox.SelectedIndex = 0;

            // FIXED: Better button sizing
            var btnOK = new Button
            {
                Text = "Select",
                DialogResult = DialogResult.OK,
                Size = new Size(140, 45), // FIXED: Larger buttons
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false
            };
            var btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Size = new Size(120, 45), // FIXED: Larger buttons
                Font = new Font(SystemFonts.DefaultFont.FontFamily, 9F),
                AutoSize = false
            };

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 90, // FIXED: Even more height for buttons
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(20, 20, 20, 20) // FIXED: More padding around buttons
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