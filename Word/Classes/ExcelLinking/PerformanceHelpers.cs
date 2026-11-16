// PerformanceHelpers.cs (.NET Framework 4.8 / C# 7.3)

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Word.ExcelLinking
{
    /// <summary>
    /// RAII wrapper that sets Excel to maximum performance mode and restores on dispose.
    /// This is critical for large workbooks with circular references.
    /// </summary>
    internal sealed class ExcelPerformanceScope : IDisposable
    {
        private readonly Excel.Application _app;
        private Excel.XlCalculation _originalCalc;
        private bool _originalScreenUpdating;
        private bool _originalDisplayAlerts;
        private bool _originalEnableEvents;
        private bool _originalAskToUpdateLinks;
        private bool _originalDisplayStatusBar;
        private bool _calcCached;
        private bool _disposed;

        public ExcelPerformanceScope(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));

            // Cache original settings - some properties can throw when no workbooks are open
            try { _originalCalc = app.Calculation; _calcCached = true; }
            catch { _originalCalc = Excel.XlCalculation.xlCalculationAutomatic; _calcCached = false; }

            try { _originalScreenUpdating = app.ScreenUpdating; } catch { _originalScreenUpdating = true; }
            try { _originalDisplayAlerts = app.DisplayAlerts; } catch { _originalDisplayAlerts = true; }
            try { _originalEnableEvents = app.EnableEvents; } catch { _originalEnableEvents = true; }
            try { _originalAskToUpdateLinks = app.AskToUpdateLinks; } catch { _originalAskToUpdateLinks = true; }
            try { _originalDisplayStatusBar = app.DisplayStatusBar; } catch { _originalDisplayStatusBar = true; }

            // Set performance mode - ORDER MATTERS
            // Disable calculation FIRST to prevent recalcs during other changes
            try { app.Calculation = Excel.XlCalculation.xlCalculationManual; } catch { }
            try { app.EnableEvents = false; } catch { }
            try { app.ScreenUpdating = false; } catch { }
            try { app.DisplayAlerts = false; } catch { }
            try { app.AskToUpdateLinks = false; } catch { }      // Suppress "Update Links" dialog
            try { app.DisplayStatusBar = false; } catch { }      // Minor perf gain
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                // Restore in reverse order
                try { _app.DisplayStatusBar = _originalDisplayStatusBar; } catch { }
                try { _app.AskToUpdateLinks = _originalAskToUpdateLinks; } catch { }
                try { _app.DisplayAlerts = _originalDisplayAlerts; } catch { }
                try { _app.ScreenUpdating = _originalScreenUpdating; } catch { }
                try { _app.EnableEvents = _originalEnableEvents; } catch { }

                // Only restore Calculation if we successfully cached it
                if (_calcCached)
                {
                    try { _app.Calculation = _originalCalc; } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error restoring Excel settings: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Clipboard helper with optimized retry logic and adaptive waiting.
    /// </summary>
    internal static class ClipboardHelper
    {
        /// <summary>
        /// Waits for clipboard to contain image data with adaptive backoff.
        /// </summary>
        /// <param name="maxWaitMs">Maximum total wait time in milliseconds</param>
        /// <param name="initialRetryMs">Initial retry interval</param>
        /// <returns>True if clipboard is ready, false if timeout</returns>
        public static bool WaitForClipboardReady(int maxWaitMs = 200, int initialRetryMs = 10)
        {
            int elapsed = 0;
            int currentWait = initialRetryMs;

            while (elapsed < maxWaitMs)
            {
                try
                {
                    // Check multiple clipboard formats
                    if (Clipboard.ContainsImage() ||
                        Clipboard.ContainsData("EnhancedMetafile") ||
                        Clipboard.ContainsData(DataFormats.MetafilePict) ||
                        Clipboard.ContainsData("Embed Source"))
                    {
                        return true;
                    }
                }
                catch
                {
                    // Clipboard access can throw if another process is using it
                }

                Thread.Sleep(currentWait);
                elapsed += currentWait;

                // Exponential backoff with cap: 10, 15, 22, 33, 50, 50, 50...
                currentWait = Math.Min(currentWait + (currentWait / 2), 50);
            }

            return false;
        }

        /// <summary>
        /// Retries a paste operation with exponential backoff.
        /// </summary>
        public static T RetryPaste<T>(Func<T> pasteAction, int maxAttempts = 3) where T : class
        {
            Exception lastEx = null;

            for (int attempt = 0; attempt < maxAttempts; attempt++)
            {
                try
                {
                    return pasteAction();
                }
                catch (COMException ex) when (attempt < maxAttempts - 1)
                {
                    lastEx = ex;
                    // Exponential backoff: 30ms, 60ms, 120ms
                    Thread.Sleep(30 * (1 << attempt));
                    System.Diagnostics.Debug.WriteLine($"Paste attempt {attempt + 1} failed, retrying...");
                }
            }

            throw lastEx ?? new Exception("Paste failed after retries");
        }

        /// <summary>
        /// Clears clipboard safely, ignoring errors.
        /// </summary>
        public static void SafeClear()
        {
            try
            {
                Clipboard.Clear();
            }
            catch
            {
                // Ignore clipboard errors
            }
        }
    }

    /// <summary>
    /// Helper methods for detecting and handling cloud/WOPI file paths.
    /// SharePoint and OneDrive files use URL-based paths that need special handling.
    /// Also handles cross-user path normalization for cloud sync folders.
    /// </summary>
    internal static class CloudPathHelper
    {
        // Placeholder used in stored paths to represent the user profile folder
        private const string UserPlaceholder = "%USERPROFILE%";

        // Known cloud sync folder names (case-insensitive matching)
        private static readonly string[] CloudSyncFolderNames = new[]
        {
            "Dropbox",
            "OneDrive",
            "OneDrive - ",      // OneDrive for Business: "OneDrive - CompanyName"
            "Google Drive",
            "Box",
            "Box Sync",
            "iCloudDrive",
            "iCloud Drive",
            "SharePoint",
            "SharePoint - "     // SharePoint sync folders
        };

        /// <summary>
        /// Checks if a path is in a user-specific cloud sync folder.
        /// </summary>
        public static bool IsUserCloudSyncPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            if (!Regex.IsMatch(path, @"^[A-Za-z]:\\Users\\[^\\]+\\", RegexOptions.IgnoreCase))
                return false;

            foreach (string folderName in CloudSyncFolderNames)
            {
                string pattern = $@"\\Users\\[^\\]+\\{Regex.Escape(folderName)}";
                if (Regex.IsMatch(path, pattern, RegexOptions.IgnoreCase))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Normalizes a path for cross-user storage by replacing the user-specific portion
        /// with a placeholder.
        /// </summary>
        public static string NormalizePath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return path;

            if (IsCloudUrl(path))
                return path;

            if (!IsUserCloudSyncPath(path))
                return path;

            var match = Regex.Match(path, @"^[A-Za-z]:\\Users\\[^\\]+\\(.+)$", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                string relativePath = match.Groups[1].Value;
                return $"{UserPlaceholder}\\{relativePath}";
            }

            return path;
        }

        /// <summary>
        /// Resolves a normalized path by replacing the placeholder with the current user's
        /// profile folder.
        /// </summary>
        public static string ResolvePath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return path;

            if (!path.StartsWith(UserPlaceholder, StringComparison.OrdinalIgnoreCase))
                return path;

            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (string.IsNullOrEmpty(userProfile))
                return path;

            return userProfile + path.Substring(UserPlaceholder.Length);
        }

        /// <summary>
        /// Checks if a path contains the user placeholder (i.e., was normalized).
        /// </summary>
        public static bool IsNormalizedPath(string path)
        {
            return !string.IsNullOrEmpty(path) &&
                   path.StartsWith(UserPlaceholder, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Checks if a path is a cloud/WOPI URL (SharePoint, OneDrive, etc.)
        /// </summary>
        public static bool IsCloudUrl(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            return path.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                   path.StartsWith("https://", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Checks if a path is any kind of cloud path (URL or local sync folder)
        /// </summary>
        public static bool IsCloudPath(string path)
        {
            return IsCloudUrl(path) || IsOneDriveSyncPath(path) || IsWopiPath(path);
        }

        /// <summary>
        /// Checks if a path appears to be a WOPI document path.
        /// </summary>
        public static bool IsWopiPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            return path.IndexOf("/_layouts/", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("/sites/", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("/personal/", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("sharepoint.com", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("1drv.ms", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("onedrive.live", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf(".sharepoint.", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// Checks if a path is a local OneDrive sync folder path.
        /// </summary>
        public static bool IsOneDriveSyncPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            return path.IndexOf("\\OneDrive\\", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("\\OneDrive -", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("/OneDrive/", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("/OneDrive -", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("\\SharePoint\\", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("\\SharePoint -", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// Extracts the filename from a path (works for both local and URL paths).
        /// </summary>
        public static string GetFileName(string path)
        {
            if (string.IsNullOrEmpty(path))
                return string.Empty;

            if (IsCloudUrl(path))
            {
                int queryIndex = path.IndexOf('?');
                if (queryIndex > 0)
                    path = path.Substring(0, queryIndex);

                int lastSlash = path.LastIndexOf('/');
                if (lastSlash >= 0 && lastSlash < path.Length - 1)
                    return Uri.UnescapeDataString(path.Substring(lastSlash + 1));
            }

            return System.IO.Path.GetFileName(path);
        }

        /// <summary>
        /// Compares two paths for equality, handling cloud path variations.
        /// </summary>
        public static bool PathsMatch(string path1, string path2)
        {
            if (string.IsNullOrEmpty(path1) || string.IsNullOrEmpty(path2))
                return false;

            if (string.Equals(path1, path2, StringComparison.OrdinalIgnoreCase))
                return true;

            string name1 = GetFileName(path1);
            string name2 = GetFileName(path2);

            if (!string.Equals(name1, name2, StringComparison.OrdinalIgnoreCase))
                return false;

            bool isCloud1 = IsCloudPath(path1);
            bool isCloud2 = IsCloudPath(path2);

            if (isCloud1 || isCloud2)
                return true;

            return false;
        }

        /// <summary>
        /// More aggressive filename-only matching for fallback scenarios.
        /// </summary>
        public static bool FileNamesMatch(string path1, string path2)
        {
            if (string.IsNullOrEmpty(path1) || string.IsNullOrEmpty(path2))
                return false;

            string name1 = GetFileName(path1);
            string name2 = GetFileName(path2);

            return string.Equals(name1, name2, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Checks if a file exists (works for both local and cloud paths).
        /// </summary>
        public static bool FileExists(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            if (IsCloudUrl(path))
                return true;

            return System.IO.File.Exists(path);
        }

        /// <summary>
        /// Gets all OneDrive and SharePoint sync folder roots on this machine.
        /// </summary>
        public static List<string> GetSyncFolderRoots()
        {
            var roots = new List<string>();

            try
            {
                string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                if (!string.IsNullOrEmpty(userProfile) && System.IO.Directory.Exists(userProfile))
                {
                    foreach (string dir in System.IO.Directory.GetDirectories(userProfile))
                    {
                        string dirName = System.IO.Path.GetFileName(dir);
                        if (dirName.StartsWith("OneDrive", StringComparison.OrdinalIgnoreCase) ||
                            dirName.StartsWith("SharePoint", StringComparison.OrdinalIgnoreCase))
                        {
                            roots.Add(dir);
                        }
                    }
                }

                string oneDrivePersonal = Environment.GetEnvironmentVariable("OneDrive");
                if (!string.IsNullOrEmpty(oneDrivePersonal) && System.IO.Directory.Exists(oneDrivePersonal))
                {
                    if (!roots.Contains(oneDrivePersonal, StringComparer.OrdinalIgnoreCase))
                        roots.Add(oneDrivePersonal);
                }

                string oneDriveBusiness = Environment.GetEnvironmentVariable("OneDriveCommercial");
                if (!string.IsNullOrEmpty(oneDriveBusiness) && System.IO.Directory.Exists(oneDriveBusiness))
                {
                    if (!roots.Contains(oneDriveBusiness, StringComparer.OrdinalIgnoreCase))
                        roots.Add(oneDriveBusiness);
                }

                string oneDriveConsumer = Environment.GetEnvironmentVariable("OneDriveConsumer");
                if (!string.IsNullOrEmpty(oneDriveConsumer) && System.IO.Directory.Exists(oneDriveConsumer))
                {
                    if (!roots.Contains(oneDriveConsumer, StringComparer.OrdinalIgnoreCase))
                        roots.Add(oneDriveConsumer);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error finding sync folders: {ex.Message}");
            }

            return roots;
        }

        /// <summary>
        /// Attempts to find a local file in OneDrive/SharePoint sync folders matching a cloud URL.
        /// </summary>
        public static string TryFindLocalSyncPath(string cloudUrl)
        {
            if (string.IsNullOrEmpty(cloudUrl))
                return null;

            string targetFileName = GetFileName(cloudUrl);
            if (string.IsNullOrEmpty(targetFileName))
                return null;

            var syncRoots = GetSyncFolderRoots();

            foreach (string root in syncRoots)
            {
                try
                {
                    string found = SearchForFile(root, targetFileName, maxDepth: 5);
                    if (found != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"Found local sync path: {cloudUrl} -> {found}");
                        return found;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error searching {root}: {ex.Message}");
                }
            }

            return null;
        }

        private static string SearchForFile(string directory, string fileName, int maxDepth, int currentDepth = 0)
        {
            if (currentDepth > maxDepth)
                return null;

            try
            {
                foreach (string file in System.IO.Directory.GetFiles(directory))
                {
                    if (string.Equals(System.IO.Path.GetFileName(file), fileName, StringComparison.OrdinalIgnoreCase))
                        return file;
                }

                foreach (string subDir in System.IO.Directory.GetDirectories(directory))
                {
                    string dirName = System.IO.Path.GetFileName(subDir);
                    if (dirName.StartsWith(".") || dirName.StartsWith("$"))
                        continue;

                    string found = SearchForFile(subDir, fileName, maxDepth, currentDepth + 1);
                    if (found != null)
                        return found;
                }
            }
            catch (UnauthorizedAccessException)
            {
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Search error in {directory}: {ex.Message}");
            }

            return null;
        }
    }

    /// <summary>
    /// Tracks workbook path updates (version upgrades or fixes) during a single batch operation.
    /// </summary>
    internal class WorkbookVersionTracker
    {
        private readonly Dictionary<string, string> pathUpdates =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public void RecordUpdate(string oldPath, string newPath)
        {
            if (!pathUpdates.ContainsKey(oldPath))
                pathUpdates[oldPath] = newPath;
        }

        public string GetUpdatedPath(string originalPath)
        {
            return pathUpdates.TryGetValue(originalPath, out string updated) ? updated : originalPath;
        }

        public bool WasUpdated(string originalPath)
        {
            return pathUpdates.ContainsKey(originalPath);
        }
    }

    /// <summary>
    /// Helper methods for finding versioned files.
    /// </summary>
    internal static class VersionHelper
    {
        public static string FindNewestVersion(string currentPath)
        {
            if (string.IsNullOrEmpty(currentPath))
                return null;

            bool currentFileExists = System.IO.File.Exists(currentPath);
            string normalizedCurrentPath = System.IO.Path.GetFullPath(currentPath);
            string directory = System.IO.Path.GetDirectoryName(normalizedCurrentPath);
            string fileName = System.IO.Path.GetFileNameWithoutExtension(normalizedCurrentPath);
            string extension = System.IO.Path.GetExtension(normalizedCurrentPath);

            var match = Regex.Match(fileName, @"^(.+?)[ _]v(\d+)$", RegexOptions.IgnoreCase);

            if (!match.Success) return null;

            string baseNameWithoutVersion = match.Groups[1].Value;
            int currentVersion = int.Parse(match.Groups[2].Value);
            string searchPattern = $"{baseNameWithoutVersion}_v*{extension}";

            try
            {
                var files = System.IO.Directory.GetFiles(directory, searchPattern);
                int highestVersion = currentFileExists ? currentVersion : -1;
                string newestFile = null;

                foreach (var file in files)
                {
                    string name = System.IO.Path.GetFileNameWithoutExtension(file);
                    var versionMatch = Regex.Match(name, @"[ _]v(\d+)$", RegexOptions.IgnoreCase);

                    if (versionMatch.Success)
                    {
                        int version = int.Parse(versionMatch.Groups[1].Value);
                        if (version > highestVersion)
                        {
                            highestVersion = version;
                            newestFile = file;
                        }
                    }
                }

                if (newestFile == null) return null;

                string normalizedNewestFile = System.IO.Path.GetFullPath(newestFile);
                if (string.Equals(normalizedNewestFile, normalizedCurrentPath, StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                return normalizedNewestFile;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error searching for versions: {ex.Message}");
                return null;
            }
        }

        public static string GetVersionDescription(string oldPath, string newPath)
        {
            string oldName = System.IO.Path.GetFileName(oldPath);
            string newName = System.IO.Path.GetFileName(newPath);
            return $"{oldName} → {newName}";
        }
    }

    /// <summary>
    /// Caches Excel workbooks during a batch operation to avoid repeated open/close.
    /// </summary>
    internal class WorkbookCache : IDisposable
    {
        private class WorkbookHandle
        {
            public Excel.Workbook Workbook { get; set; }
            public bool OpenedByUs { get; set; }
            public string ActualPath { get; set; }
            public bool WindowPrepared { get; set; }
        }

        private readonly Dictionary<string, WorkbookHandle> cache =
            new Dictionary<string, WorkbookHandle>(StringComparer.OrdinalIgnoreCase);
        private readonly Excel.Application excelApp;
        private readonly WorkbookVersionTracker versionTracker;
        private readonly bool checkForNewerVersions;

        public WorkbookCache(Excel.Application excelApp, bool checkForNewerVersions = true)
        {
            this.excelApp = excelApp;
            this.checkForNewerVersions = checkForNewerVersions;
            this.versionTracker = new WorkbookVersionTracker();
        }

        private Excel.Workbook OpenWorkbookOptimized(string path)
        {
            try
            {
                return excelApp.Workbooks.Open(
                    Filename: path,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    IgnoreReadOnlyRecommended: true,
                    AddToMru: false);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Fast open failed: {ex.Message}");

                try
                {
                    return excelApp.Workbooks.Open(
                        Filename: path,
                        UpdateLinks: 0,
                        ReadOnly: true);
                }
                catch (Exception ex2)
                {
                    System.Diagnostics.Debug.WriteLine($"Simple open failed: {ex2.Message}");
                    return excelApp.Workbooks.Open(path);
                }
            }
        }

        public Excel.Workbook GetOrOpen(string wbPath, out bool newlyOpened, out string actualPath)
        {
            newlyOpened = false;
            actualPath = wbPath;

            if (string.IsNullOrEmpty(wbPath))
                return null;

            // Resolve normalized paths
            if (CloudPathHelper.IsNormalizedPath(wbPath))
            {
                actualPath = CloudPathHelper.ResolvePath(wbPath);
                System.Diagnostics.Debug.WriteLine($"Resolved normalized path: {wbPath} -> {actualPath}");
            }

            bool isCloudPath = CloudPathHelper.IsCloudPath(actualPath);

            // Check if versioning decision already made
            if (versionTracker.WasUpdated(wbPath))
            {
                actualPath = versionTracker.GetUpdatedPath(wbPath);
            }
            else if (checkForNewerVersions && !isCloudPath)
            {
                string alternativeVersion = VersionHelper.FindNewestVersion(wbPath);
                if (alternativeVersion != null)
                {
                    bool isBrokenLink = !System.IO.File.Exists(wbPath);
                    DialogResult result;

                    if (isBrokenLink)
                    {
                        result = MessageBox.Show(
                            $"The linked workbook could not be found:\n{System.IO.Path.GetFileName(wbPath)}\n\n" +
                            $"An alternative version was found:\n{System.IO.Path.GetFileName(alternativeVersion)}\n\n" +
                            "Would you like to update all links from this missing file to the available version for this session?",
                            "Workbook Link Broken",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    }
                    else
                    {
                        string versionDesc = VersionHelper.GetVersionDescription(wbPath, alternativeVersion);
                        result = MessageBox.Show(
                            $"A newer version of the workbook was found:\n\n{versionDesc}\n\n" +
                            "Would you like to update all links to this workbook to use the newer version for this session?",
                            "Newer Workbook Version Found",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    }

                    if (result == DialogResult.Yes)
                    {
                        actualPath = alternativeVersion;
                        versionTracker.RecordUpdate(wbPath, actualPath);
                    }
                    else
                    {
                        versionTracker.RecordUpdate(wbPath, wbPath);
                    }
                }
            }

            // Check cache
            if (cache.TryGetValue(actualPath, out var handle))
            {
                return handle.Workbook;
            }
            if (!string.Equals(wbPath, actualPath, StringComparison.OrdinalIgnoreCase) &&
                cache.TryGetValue(wbPath, out handle))
            {
                return handle.Workbook;
            }

            // Check if already open in Excel
            Excel.Workbook foundByFilename = null;
            string foundByFilenamePath = null;

            foreach (Excel.Workbook openWb in excelApp.Workbooks)
            {
                try
                {
                    string openWbPath = openWb.FullName;

                    if (CloudPathHelper.PathsMatch(openWbPath, actualPath))
                    {
                        var newHandle = new WorkbookHandle
                        {
                            Workbook = openWb,
                            OpenedByUs = false,
                            ActualPath = openWbPath
                        };
                        cache[wbPath] = newHandle;
                        cache[actualPath] = newHandle;

                        if (!string.Equals(openWbPath, actualPath, StringComparison.OrdinalIgnoreCase))
                        {
                            cache[openWbPath] = newHandle;
                            System.Diagnostics.Debug.WriteLine($"Cloud workbook found: {actualPath} -> {openWbPath}");
                        }

                        actualPath = openWbPath;
                        return openWb;
                    }

                    if (foundByFilename == null && CloudPathHelper.FileNamesMatch(openWbPath, actualPath))
                    {
                        foundByFilename = openWb;
                        foundByFilenamePath = openWbPath;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error checking open workbook: {ex.Message}");
                }
            }

            // Filename-only fallback for cloud files
            if (foundByFilename != null)
            {
                bool storedPathIsCloud = CloudPathHelper.IsCloudPath(actualPath);
                bool openPathIsCloud = CloudPathHelper.IsCloudPath(foundByFilenamePath);
                bool isNormalizedPath = CloudPathHelper.IsNormalizedPath(wbPath);

                if (storedPathIsCloud || openPathIsCloud || isNormalizedPath)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"Cloud fallback: Using filename match for '{CloudPathHelper.GetFileName(actualPath)}'\n" +
                        $"  Stored path: {actualPath} (cloud: {storedPathIsCloud})\n" +
                        $"  Excel path: {foundByFilenamePath} (cloud: {openPathIsCloud})");

                    var newHandle = new WorkbookHandle
                    {
                        Workbook = foundByFilename,
                        OpenedByUs = false,
                        ActualPath = foundByFilenamePath
                    };
                    cache[wbPath] = newHandle;
                    cache[actualPath] = newHandle;
                    cache[foundByFilenamePath] = newHandle;

                    versionTracker.RecordUpdate(actualPath, foundByFilenamePath);
                    actualPath = foundByFilenamePath;
                    return foundByFilename;
                }
            }

            // Check if file exists
            if (!CloudPathHelper.FileExists(actualPath))
            {
                if (CloudPathHelper.IsCloudUrl(actualPath))
                {
                    string localPath = CloudPathHelper.TryFindLocalSyncPath(actualPath);
                    if (localPath != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"Found local sync path for cloud URL: {actualPath} -> {localPath}");
                        actualPath = localPath;
                    }
                }

                if (!CloudPathHelper.FileExists(actualPath))
                {
                    System.Diagnostics.Debug.WriteLine($"File not found: {actualPath}");
                    return null;
                }
            }

            // Open the file
            Exception lastError = null;
            string[] pathsToTry;

            if (CloudPathHelper.IsCloudUrl(actualPath))
            {
                string localSyncPath = CloudPathHelper.TryFindLocalSyncPath(actualPath);
                if (localSyncPath != null)
                    pathsToTry = new[] { actualPath, localSyncPath };
                else
                    pathsToTry = new[] { actualPath };
            }
            else
            {
                pathsToTry = new[] { actualPath };
            }

            foreach (string pathToTry in pathsToTry)
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine($"Attempting to open: {pathToTry}");
                    var wb = OpenWorkbookOptimized(pathToTry);

                    string excelReportedPath = wb.FullName;

                    var newHandle = new WorkbookHandle
                    {
                        Workbook = wb,
                        OpenedByUs = true,
                        ActualPath = excelReportedPath
                    };

                    cache[wbPath] = newHandle;
                    cache[actualPath] = newHandle;
                    cache[pathToTry] = newHandle;

                    if (!string.Equals(excelReportedPath, actualPath, StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(excelReportedPath, pathToTry, StringComparison.OrdinalIgnoreCase))
                    {
                        cache[excelReportedPath] = newHandle;
                    }

                    if (!string.Equals(pathToTry, actualPath, StringComparison.OrdinalIgnoreCase))
                    {
                        System.Diagnostics.Debug.WriteLine($"Cloud path resolved: {actualPath} -> {pathToTry} -> {excelReportedPath}");
                        versionTracker.RecordUpdate(actualPath, excelReportedPath);
                    }

                    newlyOpened = true;
                    actualPath = excelReportedPath;
                    return wb;
                }
                catch (Exception ex)
                {
                    lastError = ex;
                    System.Diagnostics.Debug.WriteLine($"Failed to open '{pathToTry}': {ex.Message}");
                }
            }

            if (CloudPathHelper.IsCloudUrl(wbPath))
            {
                System.Diagnostics.Debug.WriteLine(
                    $"Failed to open cloud workbook: {wbPath}\n" +
                    $"Error: {lastError?.Message}\n" +
                    "Tip: Try opening the Excel file manually, or ensure the file is synced locally.");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"Error opening workbook '{actualPath}': {lastError?.Message}");
            }
            return null;
        }

        public void PrepareWorkbookForCopy(Excel.Workbook wb)
        {
            if (wb == null) return;

            WorkbookHandle handle = null;
            foreach (var kvp in cache)
            {
                if (ReferenceEquals(kvp.Value.Workbook, wb))
                {
                    handle = kvp.Value;
                    break;
                }
            }

            if (handle != null && handle.WindowPrepared)
                return;

            try
            {
                if (wb.Windows.Count > 0)
                {
                    Excel.Window wbWindow = wb.Windows[1];
                    try
                    {
                        wbWindow.DisplayGridlines = false;
                        wbWindow.View = Excel.XlWindowView.xlNormalView;
                    }
                    catch { }
                    Marshal.ReleaseComObject(wbWindow);
                }
            }
            catch { }

            if (handle != null)
                handle.WindowPrepared = true;
        }

        public Excel.Workbook GetOrOpen(string wbPath, out bool newlyOpened)
        {
            return GetOrOpen(wbPath, out newlyOpened, out _);
        }

        public void Dispose()
        {
            foreach (var handle in cache.Values)
            {
                if (handle.OpenedByUs && handle.Workbook != null)
                {
                    try
                    {
                        handle.Workbook.Close(SaveChanges: false);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error closing workbook: {ex.Message}");
                    }
                }

                if (handle.Workbook != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(handle.Workbook);
                    }
                    catch { }
                }
            }
            cache.Clear();
        }
    }

    /// <summary>
    /// Collects warnings during batch operations for single summary display.
    /// </summary>
    internal class WarningCollector
    {
        private readonly List<string> warnings = new List<string>();
        private int updateCount = 0;
        private int failedCount = 0;
        private readonly int maxWarnings;

        public WarningCollector(int maxWarnings = 50)
        {
            this.maxWarnings = maxWarnings;
        }

        public void AddWarning(string warning)
        {
            if (warnings.Count < maxWarnings)
                warnings.Add(warning);
            System.Diagnostics.Debug.WriteLine($"Warning: {warning}");
        }

        public void IncrementUpdated() => updateCount++;
        public void IncrementFailed() => failedCount++;
        public int UpdateCount => updateCount;
        public int FailedCount => failedCount;
        public int WarningCount => warnings.Count;

        public string BuildSummaryMessage(int topIssues = 5)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"Updated: {updateCount}");
            sb.AppendLine($"Failed/Skipped: {failedCount}");

            if (warnings.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Top issues:");
                int show = Math.Min(topIssues, warnings.Count);
                for (int i = 0; i < show; i++)
                    sb.AppendLine("• " + warnings[i]);
                if (warnings.Count > show)
                    sb.AppendLine($"… and {warnings.Count - show} more (see debug output)");
            }

            return sb.ToString();
        }

        public void ShowSummary(string title = "Update Complete")
        {
            string message = BuildSummaryMessage(5);
            MessageBox.Show(message, title, MessageBoxButtons.OK,
                failedCount > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
        }
    }
}