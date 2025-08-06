using System;
using System.Data.SQLite;
using System.IO;

namespace PPTProductivitySuite
{
    public static class SlideLibrary
    {
        private const string LibraryFolderName = "PPTSlideLibrary";
        private const string DatabaseFileName = "SlideLibrary.db";
        private static string _libraryPath;
        private static SQLiteConnection _dbConnection;

        public static string LibraryPath => _libraryPath;
        public static SQLiteConnection DbConnection => _dbConnection;

        static SlideLibrary()
        {
            _libraryPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                LibraryFolderName);

            Directory.CreateDirectory(_libraryPath);
            InitializeDatabase();
        }

        public static void VerifyDatabase()
        {
            try
            {
                // Check if database file exists
                var dbPath = Path.Combine(_libraryPath, DatabaseFileName);
                if (!File.Exists(dbPath))
                {
                    InitializeDatabase();
                    return;
                }

                // Check if table exists
                using (var cmd = new SQLiteCommand(
                    "SELECT name FROM sqlite_master WHERE type='table' AND name='Slides'",
                    _dbConnection))
                {
                    var result = cmd.ExecuteScalar();
                    if (result == null)
                    {
                        CreateTables();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Database verification failed: {ex}");
                InitializeDatabase();
            }
        }

        private static void InitializeDatabase()
        {
            try
            {
                var dbPath = Path.Combine(_libraryPath, DatabaseFileName);
                _dbConnection = new SQLiteConnection($"Data Source={dbPath};Version=3;");
                _dbConnection.Open();
                CreateTables();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Database initialization failed: {ex}");
                throw;
            }
        }

        private static void CreateTables()
        {
            using (var cmd = new SQLiteCommand(
                @"CREATE TABLE IF NOT EXISTS Slides (
                    Id TEXT PRIMARY KEY,
                    Title TEXT NOT NULL,
                    Tags TEXT,
                    CreatedDate TEXT NOT NULL,
                    LastModified TEXT NOT NULL,
                    ThumbnailPath TEXT NOT NULL,
                    SlideFilePath TEXT NOT NULL
                )", _dbConnection))
            {
                cmd.ExecuteNonQuery();
            }
        }

        public static bool SlideExists(string title)
        {
            VerifyDatabase();

            using (var cmd = new SQLiteCommand(
                "SELECT 1 FROM Slides WHERE Title=@title",
                _dbConnection))
            {
                cmd.Parameters.AddWithValue("@title", title);
                return cmd.ExecuteScalar() != null;
            }
        }

        public static bool DeleteSlide(string slideId)
        {
            try
            {
                VerifyDatabase();

                using (var cmd = new SQLiteCommand(
                    "DELETE FROM Slides WHERE Id=@id",
                    DbConnection))
                {
                    cmd.Parameters.AddWithValue("@id", slideId);
                    int affectedRows = cmd.ExecuteNonQuery();
                    return affectedRows > 0;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to delete slide: {ex}");
                return false;
            }
        }
    }
}