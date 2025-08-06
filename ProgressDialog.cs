using System;
using System.Windows.Forms;
using System.Drawing;

namespace PPTProductivitySuite
{
    public class ProgressDialog : Form, IDisposable
    {
        private ProgressBar progressBar;
        private Label lblStatus;

        public ProgressDialog(string title, int maxSteps)
        {
            InitializeComponents();
            this.Text = title;
            this.progressBar.Maximum = maxSteps;
        }

        private void InitializeComponents()
        {
            this.progressBar = new ProgressBar();
            this.lblStatus = new Label();

            // progressBar
            this.progressBar.Dock = DockStyle.Top;
            this.progressBar.Height = 30;
            this.progressBar.Minimum = 0;

            // lblStatus
            this.lblStatus.Dock = DockStyle.Fill;
            this.lblStatus.TextAlign = ContentAlignment.MiddleCenter;

            // Form
            this.ClientSize = new Size(400, 100);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar);
        }

        public void UpdateProgress(int currentStep, string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateProgress(currentStep, message)));
                return;
            }

            this.progressBar.Value = currentStep;
            this.lblStatus.Text = message;
            Application.DoEvents();
        }
    }
}