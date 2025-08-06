using Microsoft.Office.Core;
using PPTProductivitySuite;
using System;
using System.Collections.Concurrent;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

[ComVisible(true)]
public class RibbonController : IRibbonExtensibility
{
    private IRibbonUI _ribbon;
    private static RibbonController _instance;
    private PositionData _copiedPosition;
    private Form _libraryForm;

    #region Initialization and Core Functions

    public static RibbonController Instance
    {
        get
        {
            if (_instance == null)
            {
                _instance = new RibbonController();
            }
            return _instance;
        }
    }

    public IRibbonUI Ribbon => _ribbon;

    public string GetCustomUI(string ribbonID)
    {
        try
        {
            DebugMessage($"GetCustomUI called with ribbonID: {ribbonID}", false);

            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = $"{assembly.GetName().Name}.Ribbon.xml";

            DebugMessage($"Looking for resource: {resourceName}", false);

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    var availableResources = string.Join(", ", assembly.GetManifestResourceNames());
                    DebugMessage($"Resource '{resourceName}' not found! Available resources: {availableResources}", true);
                    throw new FileNotFoundException($"Resource '{resourceName}' not found");
                }

                using (var reader = new StreamReader(stream))
                {
                    var ribbonXml = reader.ReadToEnd();
                    DebugMessage($"Full Ribbon XML loaded successfully (length: {ribbonXml.Length} chars)", false);
                    return ribbonXml;
                }
            }
        }
        catch (Exception ex)
        {
            DebugMessage($"GetCustomUI Failed: {ex}", true);
            throw;
        }
    }

    public void Ribbon_Load(IRibbonUI ribbonUI)
    {
        try
        {
            _ribbon = ribbonUI;
            DebugMessage("Ribbon_Load called successfully - ribbon should now be visible!", false);
        }
        catch (Exception ex)
        {
            DebugMessage($"Ribbon_Load failed: {ex}", true);
        }
    }

    public void InvalidateRibbon()
    {
        SafeExecute(() => _ribbon?.Invalidate(), "Ribbon invalidation");
    }

    // CRITICAL: This callback is required by Ribbon.xml
    public bool GetGroupVisible(IRibbonControl control)
    {
        // Always return true - let PowerPoint handle the visibility logic
        // This is the most reliable approach for ribbon groups
        return true;
    }

    #endregion

    public void OnTestClick(IRibbonControl control)
    {
        MessageBox.Show("Test button clicked! Ribbon is working.", "Success",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    #region Formatting Tools

    public void OnAlignClick(IRibbonControl control)
    {
        SafeExecute(() =>
        {
            var pptApp = Globals.ThisAddIn.Application;
            if (pptApp.ActiveWindow.Selection.ShapeRange.Count == 0)
            {
                MessageBox.Show("Please select shapes to align", "No Selection",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var shapes = pptApp.ActiveWindow.Selection.ShapeRange;

            switch (control.Id)
            {
                case "btnAlignLeft":
                    shapes.Align(MsoAlignCmd.msoAlignLefts, MsoTriState.msoFalse);
                    break;
                case "btnAlignCenter":
                    shapes.Align(MsoAlignCmd.msoAlignCenters, MsoTriState.msoFalse);
                    break;
                case "btnAlignRight":
                    shapes.Align(MsoAlignCmd.msoAlignRights, MsoTriState.msoFalse);
                    break;
                case "btnAlignTop":
                    shapes.Align(MsoAlignCmd.msoAlignTops, MsoTriState.msoFalse);
                    break;
                case "btnAlignMiddle":
                    shapes.Align(MsoAlignCmd.msoAlignMiddles, MsoTriState.msoFalse);
                    break;
                case "btnAlignBottom":
                    shapes.Align(MsoAlignCmd.msoAlignBottoms, MsoTriState.msoFalse);
                    break;
            }

            Debug.WriteLine($"Successfully executed alignment: {control.Id}");
        }, "Alignment");
    }

    public void OnDistributeClick(IRibbonControl control)
    {
        SafeExecute(() =>
        {
            var pptApp = Globals.ThisAddIn.Application;
            if (pptApp.ActiveWindow.Selection.ShapeRange.Count < 3)
            {
                MessageBox.Show("Select at least 3 shapes to distribute", "Insufficient Selection",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var shapes = pptApp.ActiveWindow.Selection.ShapeRange;

            if (control.Id == "btnDistributeHoriz")
            {
                shapes.Distribute(MsoDistributeCmd.msoDistributeHorizontally, MsoTriState.msoFalse);
            }
            else if (control.Id == "btnDistributeVert")
            {
                shapes.Distribute(MsoDistributeCmd.msoDistributeVertically, MsoTriState.msoFalse);
            }

            Debug.WriteLine($"Successfully executed distribution: {control.Id}");
        }, "Distribution");
    }

    #endregion

    #region Slide Library Functions

    public void OnSaveSlideClick(IRibbonControl control)
    {
        SafeExecute(() =>
        {
            SlideLibrary.VerifyDatabase();

            var pptApp = Globals.ThisAddIn.Application;
            if (pptApp.ActiveWindow.ViewType != PowerPoint.PpViewType.ppViewNormal)
            {
                MessageBox.Show("Please switch to Normal view to save slides", "View Change Required",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var currentSlide = pptApp.ActiveWindow.View.Slide as PowerPoint.Slide;
            if (currentSlide == null)
            {
                MessageBox.Show("No slide selected", "Selection Required",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var saveDialog = new SaveSlideDialog())
            {
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    SaveSlideToLibrary(currentSlide, saveDialog.SlideTitle, saveDialog.Tags);
                    MessageBox.Show("Slide saved to library!", "Success",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }, "Save Slide");
    }

    private void SaveSlideToLibrary(PowerPoint.Slide slide, string title, string[] tags)
    {
        string slideId = Guid.NewGuid().ToString();
        string slideFileName = $"{slideId}.pptx";
        string thumbFileName = $"{slideId}.png";
        string slideFilePath = Path.Combine(SlideLibrary.LibraryPath, slideFileName);
        string thumbPath = Path.Combine(SlideLibrary.LibraryPath, thumbFileName);

        // Save slide and thumbnail
        slide.Export(slideFilePath, "PPTX", 1024, 768);
        GenerateSlideThumbnail(slide, thumbPath);

        // Save to database
        using (var cmd = new SQLiteCommand(SlideLibrary.DbConnection))
        {
            cmd.CommandText = @"INSERT INTO Slides 
                              (Id, Title, Tags, CreatedDate, LastModified, ThumbnailPath, SlideFilePath)
                              VALUES (@id, @title, @tags, @created, @modified, @thumb, @slide)";

            cmd.Parameters.AddWithValue("@id", slideId);
            cmd.Parameters.AddWithValue("@title", title);
            cmd.Parameters.AddWithValue("@tags", string.Join(",", tags));
            cmd.Parameters.AddWithValue("@created", DateTime.Now.ToString("o"));
            cmd.Parameters.AddWithValue("@modified", DateTime.Now.ToString("o"));
            cmd.Parameters.AddWithValue("@thumb", thumbPath);
            cmd.Parameters.AddWithValue("@slide", slideFilePath);

            cmd.ExecuteNonQuery();
        }
    }

    public void OnShowLibraryClick(IRibbonControl control)
    {
        SafeExecute(() =>
        {
            if (_libraryForm == null || _libraryForm.IsDisposed)
            {
                _libraryForm = new SlideLibraryForm();
                _libraryForm.FormClosed += (s, e) => { _libraryForm = null; };
                _libraryForm.Show();
            }
            else
            {
                _libraryForm.BringToFront();
            }
        }, "Show Library");
    }

    #endregion

    #region Size Formatting Tools

    public void OnSameSizeClick(IRibbonControl control)
    {
        SafeExecute(() =>
        {
            var pptApp = Globals.ThisAddIn.Application;
            var selection = pptApp.ActiveWindow.Selection;

            if (selection.ShapeRange.Count < 2)
            {
                MessageBox.Show("Please select at least 2 shapes", "Selection Required",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var shapes = selection.ShapeRange;
            var referenceShape = shapes[1]; // First selected shape

            switch (control.Id)
            {
                case "btnSameWidth":
                    ApplySameWidth(shapes, referenceShape);
                    break;

                case "btnSameHeight":
                    ApplySameHeight(shapes, referenceShape);
                    break;

                case "btnSameSize":
                    ApplySameDimensions(shapes, referenceShape);
                    break;
            }

            Debug.WriteLine($"Successfully executed size adjustment: {control.Id}");
        }, "Size Adjustment");
    }

    private void ApplySameWidth(PowerPoint.ShapeRange shapes, PowerPoint.Shape reference)
    {
        foreach (PowerPoint.Shape shape in shapes)
        {
            if (shape.Id != reference.Id) // Don't modify the reference shape
            {
                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                shape.Width = reference.Width;
            }
        }
    }

    private void ApplySameHeight(PowerPoint.ShapeRange shapes, PowerPoint.Shape reference)
    {
        foreach (PowerPoint.Shape shape in shapes)
        {
            if (shape.Id != reference.Id)
            {
                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                shape.Height = reference.Height;
            }
        }
    }

    private void ApplySameDimensions(PowerPoint.ShapeRange shapes, PowerPoint.Shape reference)
    {
        foreach (PowerPoint.Shape shape in shapes)
        {
            if (shape.Id != reference.Id)
            {
                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                shape.Width = reference.Width;
                shape.Height = reference.Height;
            }
        }
    }

    #endregion

    #region Paste Functionality

    public void OnPastePlainText(IRibbonControl control)
    {
        SafeExecute(() =>
        {
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("PasteTextOnly");
        }, "Paste as Plain Text");
    }

    public void OnPasteDestinationTheme(IRibbonControl control)
    {
        SafeExecute(() =>
        {
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("PasteAndUseDestinationTheme");
        }, "Paste with Destination Theme");
    }

    public void OnPasteSourceFormatting(IRibbonControl control)
    {
        SafeExecute(() =>
        {
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("PasteSourceFormatting");
        }, "Paste with Source Formatting");
    }

    public void OnPasteImage(IRibbonControl control)
    {
        SafeExecute(() =>
        {
            var pptApp = Globals.ThisAddIn.Application;
            var slide = pptApp.ActiveWindow.View.Slide as PowerPoint.Slide;

            if (slide != null && Clipboard.ContainsImage())
            {
                var tempFile = Path.GetTempFileName() + ".png";
                Clipboard.GetImage().Save(tempFile, ImageFormat.Png);

                slide.Shapes.AddPicture(tempFile,
                    MsoTriState.msoFalse,
                    MsoTriState.msoTrue,
                    0, 0);

                File.Delete(tempFile);
            }
            else
            {
                MessageBox.Show("No image found in clipboard", "No Image",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }, "Paste Image");
    }

    #endregion

    #region Z-Order and Position Functions

    public void OnZOrderClick(IRibbonControl control)
    {
        SafeExecute(() =>
        {
            var pptApp = Globals.ThisAddIn.Application;
            if (pptApp.ActiveWindow.Selection.ShapeRange.Count == 0)
            {
                MessageBox.Show("Please select shapes to arrange", "No Selection",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var shapes = pptApp.ActiveWindow.Selection.ShapeRange;

            switch (control.Id)
            {
                case "btnBringForward":
                    shapes.ZOrder(MsoZOrderCmd.msoBringForward);
                    break;
                case "btnBringToFront":
                    shapes.ZOrder(MsoZOrderCmd.msoBringToFront);
                    break;
                case "btnSendBackward":
                    shapes.ZOrder(MsoZOrderCmd.msoSendBackward);
                    break;
                case "btnSendToBack":
                    shapes.ZOrder(MsoZOrderCmd.msoSendToBack);
                    break;
            }

            Debug.WriteLine($"Successfully executed Z-order change: {control.Id}");
        }, "Z-Order Change");
    }

    public void OnPositionClick(IRibbonControl control)
    {
        SafeExecute(() =>
        {
            var pptApp = Globals.ThisAddIn.Application;
            var selection = pptApp.ActiveWindow.Selection;

            if (selection.ShapeRange.Count == 0)
            {
                MessageBox.Show("Please select shapes", "No Selection",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            switch (control.Id)
            {
                case "btnCopyPosition":
                    var shape = selection.ShapeRange[1];
                    _copiedPosition = new PositionData
                    {
                        Left = shape.Left,
                        Top = shape.Top,
                        Width = shape.Width,
                        Height = shape.Height
                    };
                    MessageBox.Show("Position copied", "Success",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;

                case "btnPastePosition":
                    if (_copiedPosition == null)
                    {
                        MessageBox.Show("No position copied to paste", "No Position Data",
                                      MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    foreach (PowerPoint.Shape s in selection.ShapeRange)
                    {
                        s.Left = _copiedPosition.Left;
                        s.Top = _copiedPosition.Top;
                    }
                    MessageBox.Show("Position pasted", "Success",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
            }
        }, "Position Adjustment");
    }

    #endregion

    #region Helper Methods

    private void SafeExecute(Action action, string operationName)
    {
        try
        {
            action();
        }
        catch (COMException comEx)
        {
            DebugMessage($"{operationName} COM error: {comEx.Message} (Code: 0x{comEx.ErrorCode:X8})", true);
        }
        catch (Exception ex)
        {
            DebugMessage($"{operationName} failed: {ex.Message}", true);
        }
    }

    private void DebugMessage(string message, bool isError = false)
    {
        Debug.WriteLine($"[Ribbon] {message}");
        if (isError && ThisAddIn.ShowDebugMessages)
        {
            MessageBox.Show(message, "Ribbon Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void GenerateSlideThumbnail(PowerPoint.Slide slide, string outputPath)
    {
        const int thumbWidth = 200;
        const int thumbHeight = 150;

        string tempImage = Path.GetTempFileName();
        try
        {
            slide.Export(tempImage, "PNG", thumbWidth * 2, thumbHeight * 2);
            using (var original = Image.FromFile(tempImage))
            using (var thumbnail = original.GetThumbnailImage(thumbWidth, thumbHeight, null, IntPtr.Zero))
            {
                thumbnail.Save(outputPath, ImageFormat.Png);
            }
        }
        catch (Exception ex)
        {
            DebugMessage($"Thumbnail generation failed: {ex.Message}", false);
        }
        finally
        {
            if (File.Exists(tempImage))
                File.Delete(tempImage);
        }
    }

    #endregion

    // Replace your entire Mermaid section with this simpler version

    #region Mermaid Diagram Support

    private static readonly HttpClient _httpClient = new HttpClient() { Timeout = TimeSpan.FromSeconds(30) };
    private static readonly ConcurrentDictionary<string, byte[]> _mermaidCache = new ConcurrentDictionary<string, byte[]>();

    public void OnInsertMermaidClick(IRibbonControl control)
    {
        SafeExecute(() =>
        {
            var pptApp = Globals.ThisAddIn.Application;
            if (pptApp.ActiveWindow.ViewType != PowerPoint.PpViewType.ppViewNormal)
            {
                MessageBox.Show("Please switch to Normal view to insert diagrams", "View Change Required",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var slide = pptApp.ActiveWindow.View.Slide as PowerPoint.Slide;
            if (slide == null)
            {
                MessageBox.Show("No active slide found", "No Slide",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var inputDialog = new MermaidInputDialog())
            {
                if (inputDialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(inputDialog.MermaidCode))
                {
                    using (var progressDialog = new ProgressDialog("Rendering Mermaid Diagram", 4))
                    {
                        progressDialog.Show();
                        progressDialog.UpdateProgress(1, "Preparing diagram configuration...");

                        try
                        {
                            string finalMermaidCode = inputDialog.MermaidCode;

                            // Apply custom colors if requested
                            if (inputDialog.UseCustomColors && inputDialog.SelectedColorPalette != null)
                            {
                                progressDialog.UpdateProgress(2, "Applying custom color palette...");
                                var themeConfig = inputDialog.SelectedColorPalette.GenerateMermaidThemeConfig();
                                finalMermaidCode = themeConfig + "\n" + inputDialog.MermaidCode;
                            }
                            else
                            {
                                progressDialog.UpdateProgress(2, "Using default diagram colors...");
                            }

                            progressDialog.UpdateProgress(3, "Rendering diagram...");
                            var imageBytes = RenderMermaidDiagram(finalMermaidCode, progressDialog);

                            progressDialog.UpdateProgress(4, "Inserting diagram into slide...");
                            InsertImageOnSlide(slide, imageBytes);

                            MessageBox.Show("Mermaid diagram inserted successfully!", "Success",
                                          MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        finally
                        {
                            progressDialog.Close();
                        }
                    }
                }
            }
        }, "Insert Mermaid Diagram");
    }



    private byte[] RenderMermaidDiagram(string mermaidCode, ProgressDialog progressDialog = null)
    {
        // Check cache first
        if (_mermaidCache.TryGetValue(mermaidCode, out var cachedImage))
        {
            progressDialog?.UpdateProgress(3, "Using cached diagram...");
            return cachedImage;
        }

        // Try rendering methods in order of preference (highest quality first)
        var renderers = new Func<string, byte[]>[]
        {
       TryRenderWithMermaidInkAPIHighRes,
       TryRenderWithMermaidInkAPI,
       TryRenderWithKroki,
       TryRenderWithQuickChart
        };

        Exception lastException = null;
        for (int i = 0; i < renderers.Length; i++)
        {
            try
            {
                progressDialog?.UpdateProgress(3, $"Trying rendering service {i + 1}...");
                var image = renderers[i](mermaidCode);
                if (image != null && image.Length > 0)
                {
                    _mermaidCache[mermaidCode] = image;
                    return image;
                }
            }
            catch (Exception ex)
            {
                lastException = ex;
                DebugMessage($"Renderer {i + 1} failed: {ex.Message}", false);
            }
        }

        throw new InvalidOperationException($"All Mermaid rendering methods failed. Last error: {lastException?.Message}");
    }

    #region Rendering Methods

    private byte[] TryRenderWithMermaidInkAPIHighRes(string mermaidCode)
    {
        try
        {
            // Use very high resolution PNG directly (no SVG conversion needed)
            var encodedCode = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(mermaidCode));
            var url = $"https://mermaid.ink/img/{encodedCode}?type=png&theme=base&width=2400&height=1800&scale=3";

            var response = _httpClient.GetAsync(url).Result;
            response.EnsureSuccessStatusCode();

            var imageBytes = response.Content.ReadAsByteArrayAsync().Result;

            // Validate that we got a valid image
            if (imageBytes.Length < 100)
                throw new InvalidOperationException("Invalid image data received");

            return imageBytes;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Mermaid.ink High-Res API failed: {ex.Message}", ex);
        }
    }

    private byte[] TryRenderWithMermaidInkAPI(string mermaidCode)
    {
        try
        {
            // Encode the mermaid code
            var encodedCode = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(mermaidCode));
            // Request higher resolution than before: 1920x1440 (4:3 ratio) with base theme
            var url = $"https://mermaid.ink/img/{encodedCode}?type=png&theme=base&width=1920&height=1440&scale=2";

            var response = _httpClient.GetAsync(url).Result;
            response.EnsureSuccessStatusCode();

            var imageBytes = response.Content.ReadAsByteArrayAsync().Result;

            // Validate that we got a valid image
            if (imageBytes.Length < 100) // Too small to be a valid image
                throw new InvalidOperationException("Invalid image data received");

            return imageBytes;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Mermaid.ink API failed: {ex.Message}", ex);
        }
    }

    private byte[] TryRenderWithKroki(string mermaidCode)
    {
        try
        {
            var compressed = CompressString(mermaidCode);
            var encoded = Convert.ToBase64String(compressed)
                .Replace('+', '-')
                .Replace('/', '_')
                .TrimEnd('=');

            // Note: Kroki doesn't support size parameters in URL, but typically renders at good resolution
            var url = $"https://kroki.io/mermaid/png/{encoded}";

            var response = _httpClient.GetAsync(url).Result;
            response.EnsureSuccessStatusCode();

            return response.Content.ReadAsByteArrayAsync().Result;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Kroki API failed: {ex.Message}", ex);
        }
    }

    private byte[] TryRenderWithQuickChart(string mermaidCode)
    {
        try
        {
            // Request higher resolution: 1920x1440 with high DPI
            var jsonPayload = "{\"chart\":\"" + EscapeJsonString(mermaidCode) + "\",\"format\":\"png\",\"width\":1920,\"height\":1440,\"devicePixelRatio\":2}";
            var content = new StringContent(jsonPayload, System.Text.Encoding.UTF8, "application/json");

            var response = _httpClient.PostAsync("https://quickchart.io/chart", content).Result;
            response.EnsureSuccessStatusCode();

            return response.Content.ReadAsByteArrayAsync().Result;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"QuickChart API failed: {ex.Message}", ex);
        }
    }

    #endregion

    #region Helper Methods

    private void InsertImageOnSlide(PowerPoint.Slide slide, byte[] imageBytes)
    {
        string tempImagePath = null;
        try
        {
            // Create temporary file
            tempImagePath = Path.Combine(Path.GetTempPath(), $"mermaid_{Guid.NewGuid()}.png");
            File.WriteAllBytes(tempImagePath, imageBytes);

            // Insert image into PowerPoint
            var shape = slide.Shapes.AddPicture(
                tempImagePath,
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoTrue,
                50, 50, // Default position
                -1, -1  // Use original size
            );

            // Center the image on the slide
            var slideWidth = slide.Master.Width;
            var slideHeight = slide.Master.Height;

            shape.Left = (slideWidth - shape.Width) / 2;
            shape.Top = (slideHeight - shape.Height) / 2;

            DebugMessage("Mermaid diagram inserted successfully", false);
        }
        finally
        {
            // Clean up temporary file
            if (tempImagePath != null && File.Exists(tempImagePath))
            {
                try
                {
                    File.Delete(tempImagePath);
                }
                catch (Exception ex)
                {
                    DebugMessage($"Failed to delete temp file: {ex.Message}", false);
                }
            }
        }
    }

    private string EscapeJsonString(string str)
    {
        return str.Replace("\\", "\\\\")
                  .Replace("\"", "\\\"")
                  .Replace("\r", "\\r")
                  .Replace("\n", "\\n")
                  .Replace("\t", "\\t");
    }

    private byte[] CompressString(string text)
    {
        var bytes = System.Text.Encoding.UTF8.GetBytes(text);
        using (var msi = new MemoryStream(bytes))
        using (var mso = new MemoryStream())
        {
            using (var gs = new System.IO.Compression.GZipStream(mso, System.IO.Compression.CompressionMode.Compress))
            {
                msi.CopyTo(gs);
            }
            return mso.ToArray();
        }
    }

    #endregion

    #endregion

    private class PositionData
    {
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
    }
}

