using System;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

namespace PPTProductivitySuite
{
    public class ThemeCustomizationDialog : Form
    {
        public ThemeColorManager.MermaidColorMapping CustomMapping { get; private set; }

        private readonly string[] _colorOptions = {
            "Background1", "Text1", "Background2", "Text2",
            "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6",
            "Hyperlink", "FollowedHyperlink"
        };

        private TableLayoutPanel _layoutPanel;
        private ComboBox[] _colorCombos;
        private Label[] _colorLabels;
        private Button _btnOK;
        private Button _btnCancel;
        private Button _btnReset;

        public ThemeCustomizationDialog(ThemeColorManager.MermaidColorMapping currentMapping)
        {
            CustomMapping = new ThemeColorManager.MermaidColorMapping
            {
                PrimaryColor = currentMapping.PrimaryColor,
                SecondaryColor = currentMapping.SecondaryColor,
                TertiaryColor = currentMapping.TertiaryColor,
                PrimaryTextColor = currentMapping.PrimaryTextColor,
                SecondaryTextColor = currentMapping.SecondaryTextColor,
                PrimaryBorderColor = currentMapping.PrimaryBorderColor,
                LineColor = currentMapping.LineColor,
                Background = currentMapping.Background,
                MainBkg = currentMapping.MainBkg
            };

            InitializeComponents();
            LoadCurrentMapping();
        }

        private void InitializeComponents()
        {
            Text = "Customize Theme Color Mapping";
            Size = new Size(500, 400);
            StartPosition = FormStartPosition.CenterParent;
            ShowIcon = false;
            MaximizeBox = false;
            MinimizeBox = false;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            // Create layout panel
            _layoutPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 10,
                Padding = new Padding(10),
                AutoScroll = true
            };

            // Set column styles
            _layoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            _layoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            var mappingProperties = new[]
            {
                ("Primary Color:", nameof(CustomMapping.PrimaryColor)),
                ("Secondary Color:", nameof(CustomMapping.SecondaryColor)),
                ("Tertiary Color:", nameof(CustomMapping.TertiaryColor)),
                ("Primary Text Color:", nameof(CustomMapping.PrimaryTextColor)),
                ("Secondary Text Color:", nameof(CustomMapping.SecondaryTextColor)),
                ("Primary Border Color:", nameof(CustomMapping.PrimaryBorderColor)),
                ("Line Color:", nameof(CustomMapping.LineColor)),
                ("Background:", nameof(CustomMapping.Background)),
                ("Main Background:", nameof(CustomMapping.MainBkg))
            };

            _colorLabels = new Label[mappingProperties.Length];
            _colorCombos = new ComboBox[mappingProperties.Length];

            for (int i = 0; i < mappingProperties.Length; i++)
            {
                // Label
                _colorLabels[i] = new Label
                {
                    Text = mappingProperties[i].Item1,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right,
                    Margin = new Padding(3)
                };

                // ComboBox
                _colorCombos[i] = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right,
                    Margin = new Padding(3)
                };
                _colorCombos[i].Items.AddRange(_colorOptions);

                _layoutPanel.Controls.Add(_colorLabels[i], 0, i);
                _layoutPanel.Controls.Add(_colorCombos[i], 1, i);
            }

            // Button panel
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(10)
            };

            _btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Size = new Size(80, 30),
                UseVisualStyleBackColor = true
            };

            _btnOK = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Size = new Size(80, 30),
                UseVisualStyleBackColor = true
            };

            _btnReset = new Button
            {
                Text = "Reset to Corporate",
                Size = new Size(120, 30),
                UseVisualStyleBackColor = true
            };

            buttonPanel.Controls.Add(_btnCancel);
            buttonPanel.Controls.Add(_btnOK);
            buttonPanel.Controls.Add(_btnReset);

            Controls.Add(_layoutPanel);
            Controls.Add(buttonPanel);

            AcceptButton = _btnOK;
            CancelButton = _btnCancel;

            // Event handlers
            _btnOK.Click += BtnOK_Click;
            _btnReset.Click += BtnReset_Click;
        }

        private void LoadCurrentMapping()
        {
            _colorCombos[0].Text = CustomMapping.PrimaryColor;
            _colorCombos[1].Text = CustomMapping.SecondaryColor;
            _colorCombos[2].Text = CustomMapping.TertiaryColor;
            _colorCombos[3].Text = CustomMapping.PrimaryTextColor;
            _colorCombos[4].Text = CustomMapping.SecondaryTextColor;
            _colorCombos[5].Text = CustomMapping.PrimaryBorderColor;
            _colorCombos[6].Text = CustomMapping.LineColor;
            _colorCombos[7].Text = CustomMapping.Background;
            _colorCombos[8].Text = CustomMapping.MainBkg;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            CustomMapping.PrimaryColor = _colorCombos[0].Text;
            CustomMapping.SecondaryColor = _colorCombos[1].Text;
            CustomMapping.TertiaryColor = _colorCombos[2].Text;
            CustomMapping.PrimaryTextColor = _colorCombos[3].Text;
            CustomMapping.SecondaryTextColor = _colorCombos[4].Text;
            CustomMapping.PrimaryBorderColor = _colorCombos[5].Text;
            CustomMapping.LineColor = _colorCombos[6].Text;
            CustomMapping.Background = _colorCombos[7].Text;
            CustomMapping.MainBkg = _colorCombos[8].Text;
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            CustomMapping = ThemeColorManager.MermaidColorMapping.Presets["Corporate"];
            LoadCurrentMapping();
        }
    }
}