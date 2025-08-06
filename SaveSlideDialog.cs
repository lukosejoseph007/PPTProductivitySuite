using System;
using System.Linq;
using System.Windows.Forms;

namespace PPTProductivitySuite
{
    public partial class SaveSlideDialog : Form
    {
        public string SlideTitle => txtTitle.Text;
        public string[] Tags => txtTags.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(t => t.Trim())
                                .ToArray();

        private TextBox txtTitle;
        private TextBox txtTags;
        private Button btnSave;
        private Button btnCancel;
        private Label lblTitle;
        private Label lblTags;
        public bool AllowMultiple { get; set; } = false;
        public SaveSlideDialog()
        {
            InitializeComponents();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!AllowMultiple && SlideLibrary.SlideExists(txtTitle.Text))
            {
                MessageBox.Show("A slide with this title already exists",
                              "Duplicate",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Warning);
                return;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void InitializeComponents()
        {
            this.lblTitle = new Label();
            this.txtTitle = new TextBox();
            this.lblTags = new Label();
            this.txtTags = new TextBox();
            this.btnSave = new Button();
            this.btnCancel = new Button();

            // lblTitle
            this.lblTitle.AutoSize = true;
            this.lblTitle.Location = new System.Drawing.Point(20, 20);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Text = "Slide Title:";

            // txtTitle
            this.txtTitle.Location = new System.Drawing.Point(120, 20);
            this.txtTitle.Name = "txtTitle";
            this.txtTitle.Width = 250;

            // lblTags
            this.lblTags.AutoSize = true;
            this.lblTags.Location = new System.Drawing.Point(20, 60);
            this.lblTags.Name = "lblTags";
            this.lblTags.Text = "Tags (comma separated):";

            // txtTags
            this.txtTags.Location = new System.Drawing.Point(120, 60);
            this.txtTags.Name = "txtTags";
            this.txtTags.Width = 250;

            // btnSave
            this.btnSave.Text = "Save";
            this.btnSave.DialogResult = DialogResult.OK;
            this.btnSave.Location = new System.Drawing.Point(200, 120);
            this.btnSave.Size = new System.Drawing.Size(80, 30);

            // btnCancel
            this.btnCancel.Text = "Cancel";
            this.btnCancel.DialogResult = DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(290, 120);
            this.btnCancel.Size = new System.Drawing.Size(80, 30);

            // Form
            this.Text = "Save Slide to Library";
            this.ClientSize = new System.Drawing.Size(400, 200);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.AcceptButton = btnSave;
            this.CancelButton = btnCancel;

            this.Controls.Add(lblTitle);
            this.Controls.Add(txtTitle);
            this.Controls.Add(lblTags);
            this.Controls.Add(txtTags);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnCancel);
        }
    }
}