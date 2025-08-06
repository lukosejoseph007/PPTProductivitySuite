using Microsoft.Office.Core;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using System.Threading;

namespace PPTProductivitySuite
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("A69FE362-8140-4345-82FB-A4CFFE0ABCF1")]
    public partial class ThisAddIn
    {
        private static readonly string LogFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "PPTAddInLog.txt");

        public static bool ShowDebugMessages { get; set; } = true;
        private const string AddInRegistryPath = @"Software\Microsoft\Office\PowerPoint\Addins\PPTProductivitySuite";
        private static readonly Mutex DatabaseMutex = new Mutex(false, "PPTProductivitySuite_DB_Mutex");

        [ComRegisterFunction]
        public static void Register(Type t)
        {
            try
            {
                using (var clsidKey = Registry.ClassesRoot.CreateSubKey($@"CLSID\{{{t.GUID}}}"))
                {
                    clsidKey.SetValue("", "PPT Productivity Suite");

                    using (var inprocKey = clsidKey.CreateSubKey("InprocServer32"))
                    {
                        inprocKey.SetValue("", "mscoree.dll");
                        inprocKey.SetValue("ThreadingModel", "Both");
                        inprocKey.SetValue("Class", t.FullName);
                        inprocKey.SetValue("Assembly", t.Assembly.FullName);
                        inprocKey.SetValue("RuntimeVersion", "v4.0.30319");
                    }

                    clsidKey.CreateSubKey("Programmable");
                }

                using (var addinKey = Registry.CurrentUser.CreateSubKey(AddInRegistryPath))
                {
                    addinKey.SetValue("Description", "PPT Productivity Suite");
                    addinKey.SetValue("FriendlyName", "PPT Productivity Suite");
                    addinKey.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                    addinKey.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
                }

                DebugStatic("COM registration completed successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"COM registration failed: {ex}",
                    "Registration Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                throw;
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
            try
            {
                Registry.ClassesRoot.DeleteSubKeyTree($@"CLSID\{{{t.GUID}}}");
                Registry.CurrentUser.DeleteSubKeyTree(AddInRegistryPath);
                DebugStatic("COM unregistration completed successfully");
            }
            catch (Exception ex)
            {
                DebugStatic($"Unregistration failed: {ex}", true);
            }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            System.Windows.Forms.Application.ThreadException += Application_ThreadException;

            DebugLog("Add-in startup initiated");

            // Initialize database with mutex protection
            DatabaseMutex.WaitOne();
            try
            {
                SlideLibrary.VerifyDatabase();
                DebugLog("Database verified successfully");
            }
            catch (Exception ex)
            {
                DebugLog($"Database initialization failed: {ex}", true);
                MessageBox.Show("Failed to initialize slide library. Some features may not work.",
                    "Initialization Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            finally
            {
                DatabaseMutex.ReleaseMutex();
            }

            CleanupTemporaryFiles();

            // FIXED: Add selection change handler without immediate ribbon invalidation
            Application.WindowSelectionChange += Application_WindowSelectionChange;

            VerifyAddInRegistration();
            DebugLog("Add-in startup completed successfully");
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection selection)
        {
            try
            {
                // Only invalidate ribbon if it's been loaded and after a small delay
                // This prevents issues during PowerPoint startup
                var timer = new System.Windows.Forms.Timer();
                timer.Interval = 100; // 100ms delay
                timer.Tick += (sender, e) =>
                {
                    timer.Stop();
                    timer.Dispose();

                    try
                    {
                        var ribbon = RibbonController.Instance?.Ribbon;
                        if (ribbon != null)
                        {
                            ribbon.Invalidate();
                            DebugLog("Ribbon invalidated on selection change");
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLog($"Ribbon invalidation failed: {ex.Message}");
                    }
                };
                timer.Start();
            }
            catch (Exception ex)
            {
                DebugLog($"Selection change handler failed: {ex.Message}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                DebugLog("Add-in shutdown initiated");

                // Remove event handlers
                Application.WindowSelectionChange -= Application_WindowSelectionChange;

                // Properly close database connection with mutex protection
                DatabaseMutex.WaitOne();
                try
                {
                    if (SlideLibrary.DbConnection?.State == System.Data.ConnectionState.Open)
                    {
                        SlideLibrary.DbConnection.Close();
                        DebugLog("Database connection closed");
                    }
                }
                finally
                {
                    DatabaseMutex.ReleaseMutex();
                }

                // Clean up COM objects
                while (Marshal.ReleaseComObject(Application) > 0) { }
                DebugLog("COM objects released");
            }
            catch (Exception ex)
            {
                DebugLog($"Error during shutdown: {ex}", true);
            }
            finally
            {
                DebugLog("Add-in shutdown completed");
            }
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var ex = e.ExceptionObject as Exception;
            DebugLog($"Unhandled exception: {ex?.Message ?? "Unknown error"}", true);
        }

        private void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            DebugLog($"Thread exception: {e.Exception.Message}", true);
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            DebugLog("Creating ribbon extensibility object");
            return RibbonController.Instance; // Use singleton instance
        }

        private void VerifyAddInRegistration()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(AddInRegistryPath))
                {
                    if (key == null)
                    {
                        DebugLog("Registry key not found", true);
                        return;
                    }

                    var loadBehavior = key.GetValue("LoadBehavior")?.ToString() ?? "null";
                    DebugLog($"Registry verification passed. LoadBehavior: {loadBehavior}");
                }

                using (var key = Registry.ClassesRoot.OpenSubKey(
                    $@"CLSID\\{{{GetType().GUID}}}\\InprocServer32"))
                {
                    if (key == null)
                    {
                        DebugLog("COM registration missing", true);
                    }
                    else
                    {
                        DebugLog("COM registration verification passed");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLog($"Registration check failed: {ex.Message}", true);
            }
        }

        private void DebugLog(string message, bool isError = false)
        {
            try
            {
                var logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {(isError ? "ERROR: " : "")}{message}";
                Debug.WriteLine(logMessage);
                File.AppendAllText(LogFilePath, logMessage + Environment.NewLine);
            }
            catch { /* Prevent logging failures from crashing the add-in */ }
        }

        private static void DebugStatic(string message, bool isError = false)
        {
            try
            {
                var logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {(isError ? "ERROR: " : "")}{message}";
                Debug.WriteLine(logMessage);
                File.AppendAllText(Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "PPTAddInLog.txt"), logMessage + Environment.NewLine);
            }
            catch { /* Prevent logging failures from crashing the add-in */ }
        }

        private void CleanupTemporaryFiles()
        {
            try
            {
                var tempPath = Path.GetTempPath();
                var tempFiles = Directory.GetFiles(tempPath, "tmp*.png");

                foreach (var file in tempFiles)
                {
                    try
                    {
                        if (File.GetCreationTime(file) < DateTime.Now.AddDays(-1))
                        {
                            File.Delete(file);
                            DebugLog($"Deleted old temp file: {Path.GetFileName(file)}");
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLog($"Failed to delete temp file {Path.GetFileName(file)}: {ex.Message}");
                    }
                }
                DebugLog("Temporary files cleanup completed");
            }
            catch (Exception ex)
            {
                DebugLog($"Temp file cleanup failed: {ex}", true);
            }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}