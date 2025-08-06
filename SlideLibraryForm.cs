using Microsoft.Office.Core;
using System;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTProductivitySuite
{
    public partial class SlideLibraryForm : Form
    {
        private FlowLayoutPanel flowPanel;
        private TextBox searchBox;

        public SlideLibraryForm()
        {
            InitializeComponents();
            LoadLibraryItems();
        }

        private void InitializeComponents()
        {
            this.searchBox = new TextBox();
            this.flowPanel = new FlowLayoutPanel();

            // searchBox
            this.searchBox.Dock = DockStyle.Top;
            this.searchBox.Margin = new Padding(10);
            this.searchBox.Text = "Search slides...";
            this.searchBox.ForeColor = Color.Gray;
            this.searchBox.GotFocus += (s, e) =>
            {
                if (searchBox.Text == "Search slides...")
                {
                    searchBox.Text = "";
                    searchBox.ForeColor = Color.Black;
                }
            };
            this.searchBox.LostFocus += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(searchBox.Text))
                {
                    searchBox.Text = "Search slides...";
                    searchBox.ForeColor = Color.Gray;
                }
            };
            this.searchBox.TextChanged += (s, e) => FilterItems(searchBox.Text);

            // flowPanel
            this.flowPanel.Dock = DockStyle.Fill;
            this.flowPanel.AutoScroll = true;
            this.flowPanel.WrapContents = true;
            this.flowPanel.Padding = new Padding(10);

            // Form
            this.Text = "Slide Library";
            this.ClientSize = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            this.Controls.Add(flowPanel);
            this.Controls.Add(searchBox);
        }

        public void RefreshLibrary()
        {
            flowPanel.Controls.Clear();
            LoadLibraryItems();
        }

        private void LoadLibraryItems()
        {
            try
            {
                SlideLibrary.VerifyDatabase();

                using (var cmd = new SQLiteCommand(SlideLibrary.DbConnection))
                {
                    cmd.CommandText = "SELECT * FROM Slides ORDER BY LastModified DESC";
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var item = new SlideLibraryItem
                            {
                                Id = reader["Id"].ToString(),
                                Title = reader["Title"].ToString(),
                                Tags = reader["Tags"].ToString().Split(','),
                                CreatedDate = DateTime.Parse(reader["CreatedDate"].ToString()),
                                LastModified = DateTime.Parse(reader["LastModified"].ToString()),
                                ThumbnailPath = reader["ThumbnailPath"].ToString(),
                                SlideFilePath = reader["SlideFilePath"].ToString()
                            };

                            AddLibraryItemToPanel(item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load library: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddLibraryItemToPanel(SlideLibraryItem item)
        {
            var panel = new Panel
            {
                Width = 220,
                Height = 200,
                Margin = new Padding(10),
                BorderStyle = BorderStyle.FixedSingle,
                Tag = item
            };

            // Thumbnail
            var pictureBox = new PictureBox
            {
                Width = 200,
                Height = 120,
                Location = new Point(10, 10),
                SizeMode = PictureBoxSizeMode.Zoom,
                Image = File.Exists(item.ThumbnailPath) ? Image.FromFile(item.ThumbnailPath) : null,
                Cursor = Cursors.Hand
            };
            pictureBox.Click += (s, e) => InsertSlide(item);
            panel.Controls.Add(pictureBox);

            // Title
            var titleLabel = new Label
            {
                Text = item.Title,
                Location = new Point(10, 140),
                Width = 200,
                Font = new Font(Font, FontStyle.Bold)
            };
            panel.Controls.Add(titleLabel);

            // Tags
            var tagsLabel = new Label
            {
                Text = string.Join(", ", item.Tags),
                Location = new Point(10, 160),
                Width = 200,
                ForeColor = Color.Gray
            };
            panel.Controls.Add(tagsLabel);

            // Button container
            var buttonPanel = new Panel
            {
                Location = new Point(10, 180),
                Size = new Size(200, 25)
            };

            // Insert button
            var insertBtn = new Button
            {
                Text = "Insert",
                Location = new Point(0, 0),
                Size = new Size(80, 25)
            };
            insertBtn.Click += (s, e) => InsertSlide(item);
            buttonPanel.Controls.Add(insertBtn);

            // Delete button
            var deleteBtn = new Button
            {
                Text = "Delete",
                Location = new Point(90, 0),
                Size = new Size(80, 25),
                BackColor = Color.LightCoral,
                FlatStyle = FlatStyle.Flat
            };
            deleteBtn.FlatAppearance.BorderSize = 0;
            deleteBtn.Click += (s, e) => DeleteSlide(item, panel);
            buttonPanel.Controls.Add(deleteBtn);

            panel.Controls.Add(buttonPanel);
            flowPanel.Controls.Add(panel);
        }

        private void InsertSlide(SlideLibraryItem item)
        {
            try
            {
                var pptApp = Globals.ThisAddIn.Application;
                var activePresentation = pptApp.ActivePresentation;
                var currentWindow = pptApp.ActiveWindow;

                // Open the slide file (hidden)
                var libPresentation = pptApp.Presentations.Open(item.SlideFilePath,
                    MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                // Copy the slide
                libPresentation.Slides[1].Copy();

                // Paste into current presentation
                var newSlide = activePresentation.Slides.Paste(currentWindow.View.Slide.SlideIndex + 1);

                // Close the library presentation without saving
                libPresentation.Close();

                // Select the new slide
                currentWindow.View.GotoSlide(newSlide.SlideIndex);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to insert slide: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteSlide(SlideLibraryItem item, Panel panel)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                if (MessageBox.Show($"Delete '{item.Title}' from library?",
                                  "Confirm Delete",
                                  MessageBoxButtons.YesNo,
                                  MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    // Delete from database
                    bool success = SlideLibrary.DeleteSlide(item.Id);

                    if (success)
                    {
                        // Delete associated files
                        try { File.Delete(item.SlideFilePath); } catch { /* Ignore */ }
                        try { File.Delete(item.ThumbnailPath); } catch { /* Ignore */ }

                        // Remove from UI
                        flowPanel.Controls.Remove(panel);
                        panel.Dispose();

                        MessageBox.Show("Slide deleted successfully",
                                      "Success",
                                      MessageBoxButtons.OK,
                                      MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Failed to delete slide from library",
                                      "Error",
                                      MessageBoxButtons.OK,
                                      MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error deleting slide: {ex.Message}",
                              "Error",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Error);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void FilterItems(string searchText)
        {
            foreach (Panel panel in flowPanel.Controls)
            {
                var item = panel.Tag as SlideLibraryItem;
                bool matches = string.IsNullOrEmpty(searchText) ||
                    item.Title.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    item.Tags.Any(t => t.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0);

                panel.Visible = matches;
            }
        }
    }
}