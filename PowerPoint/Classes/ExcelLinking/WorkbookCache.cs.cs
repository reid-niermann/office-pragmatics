// WorkbookCache.cs (.NET Framework 4.8 / C# 7.3)

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_PowerPoint.ExcelLinking
{
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
        /// These paths contain the username and need normalization for cross-user compatibility.
        /// Examples:
        ///   C:\Users\john\Dropbox\...
        ///   C:\Users\jane\OneDrive - Company\...
        ///   C:\Users\bob\Google Drive\...
        /// </summary>
        public static bool IsUserCloudSyncPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            // Check if path is under a Users folder
            // Pattern: [Drive]:\Users\[username]\[CloudFolder]\...
            if (!Regex.IsMatch(path, @"^[A-Za-z]:\\Users\\[^\\]+\\", RegexOptions.IgnoreCase))
                return false;

            // Check if it contains a known cloud sync folder
            foreach (string folderName in CloudSyncFolderNames)
            {
                // Match pattern like \Dropbox\ or \OneDrive - Company\
                string pattern = $@"\\Users\\[^\\]+\\{Regex.Escape(folderName)}";
                if (Regex.IsMatch(path, pattern, RegexOptions.IgnoreCase))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Normalizes a path for cross-user storage by replacing the user-specific portion
        /// with a placeholder. This allows the path to work on different computers.
        /// 
        /// Input:  C:\Users\john\Dropbox\Projects\file.xlsx
        /// Output: %USERPROFILE%\Dropbox\Projects\file.xlsx
        /// </summary>
        public static string NormalizePath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return path;

            // Don't normalize URLs
            if (IsCloudUrl(path))
                return path;

            // Check if this is a user-specific cloud sync path
            if (!IsUserCloudSyncPath(path))
                return path;

            // Extract the portion after C:\Users\username\
            // Pattern: [Drive]:\Users\[username]\[rest of path]
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
        /// profile folder. This allows paths stored on one computer to work on another.
        /// 
        /// Input:  %USERPROFILE%\Dropbox\Projects\file.xlsx
        /// Output: C:\Users\currentuser\Dropbox\Projects\file.xlsx
        /// </summary>
        public static string ResolvePath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return path;

            // Check if path contains our placeholder
            if (!path.StartsWith(UserPlaceholder, StringComparison.OrdinalIgnoreCase))
                return path;

            // Get current user's profile folder
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (string.IsNullOrEmpty(userProfile))
                return path;

            // Replace placeholder with actual path
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
        /// WOPI paths can have special formats beyond standard URLs.
        /// </summary>
        public static bool IsWopiPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            // WOPI document IDs and special paths
            // Some WOPI paths contain document IDs or special markers
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
        /// These are local paths but represent cloud-synced files.
        /// Examples:
        ///   C:\Users\username\OneDrive - Company Name\...
        ///   C:\Users\username\OneDrive\...
        ///   D:\OneDrive - Personal\...
        /// </summary>
        public static bool IsOneDriveSyncPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            // Check for common OneDrive sync folder patterns
            return path.IndexOf("\\OneDrive\\", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("\\OneDrive -", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("/OneDrive/", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   path.IndexOf("/OneDrive -", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   // Also check for SharePoint sync folders
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

            // For URLs, handle query strings and path separators
            if (IsCloudUrl(path))
            {
                // Remove query string if present
                int queryIndex = path.IndexOf('?');
                if (queryIndex > 0)
                    path = path.Substring(0, queryIndex);

                // Get last segment
                int lastSlash = path.LastIndexOf('/');
                if (lastSlash >= 0 && lastSlash < path.Length - 1)
                    return Uri.UnescapeDataString(path.Substring(lastSlash + 1));
            }

            return System.IO.Path.GetFileName(path);
        }

        /// <summary>
        /// Compares two paths for equality, handling cloud path variations.
        /// Cloud paths may have slight differences in encoding or format.
        /// </summary>
        public static bool PathsMatch(string path1, string path2)
        {
            if (string.IsNullOrEmpty(path1) || string.IsNullOrEmpty(path2))
                return false;

            // Direct comparison first
            if (string.Equals(path1, path2, StringComparison.OrdinalIgnoreCase))
                return true;

            string name1 = GetFileName(path1);
            string name2 = GetFileName(path2);

            // If filenames don't match, paths definitely don't match
            if (!string.Equals(name1, name2, StringComparison.OrdinalIgnoreCase))
                return false;

            // Filenames match - check if paths are related cloud/sync paths
            bool isCloud1 = IsCloudPath(path1);
            bool isCloud2 = IsCloudPath(path2);

            // If EITHER path is cloud-related, and filename matches, consider same
            // This handles the case where stored path and Excel path have different formats
            if (isCloud1 || isCloud2)
                return true;

            // Both are regular local paths with same filename but different folders
            // This is NOT a match (could be different files with same name)
            return false;
        }

        /// <summary>
        /// More aggressive filename-only matching for fallback scenarios.
        /// Use this when path matching fails but we want to find any open workbook with same name.
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
        /// For cloud paths, we can't easily check existence, so we return true
        /// and let the Open call handle errors.
        /// </summary>
        public static bool FileExists(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            // Cloud URLs - assume they exist, let Open handle errors
            if (IsCloudUrl(path))
                return true;

            // Local paths (including OneDrive sync folders) - check normally
            return System.IO.File.Exists(path);
        }

        /// <summary>
        /// Gets all OneDrive and SharePoint sync folder roots on this machine.
        /// Searches common locations and environment variables.
        /// </summary>
        public static List<string> GetSyncFolderRoots()
        {
            var roots = new List<string>();

            try
            {
                // Check user profile folder for OneDrive folders
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

                // Check OneDrive environment variables
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
        /// Returns null if not found.
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
                    // Search for the file in sync folders (limit depth to avoid long searches)
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

        /// <summary>
        /// Searches for a file by name in a directory tree.
        /// </summary>
        private static string SearchForFile(string directory, string fileName, int maxDepth, int currentDepth = 0)
        {
            if (currentDepth > maxDepth)
                return null;

            try
            {
                // Check files in current directory
                foreach (string file in System.IO.Directory.GetFiles(directory))
                {
                    if (string.Equals(System.IO.Path.GetFileName(file), fileName, StringComparison.OrdinalIgnoreCase))
                        return file;
                }

                // Recurse into subdirectories
                foreach (string subDir in System.IO.Directory.GetDirectories(directory))
                {
                    // Skip hidden and system folders
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
                // Skip folders we can't access
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
    /// This is the key to preventing repeated popups.
    /// </summary>
    internal class WorkbookVersionTracker
    {
        private readonly Dictionary<string, string> pathUpdates =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Records that an old path should be replaced with a new path for the remainder of the session.
        /// </summary>
        public void RecordUpdate(string oldPath, string newPath)
        {
            if (!pathUpdates.ContainsKey(oldPath))
                pathUpdates[oldPath] = newPath;
        }

        /// <summary>
        /// Gets the updated path for a workbook, or returns the original if no update has been recorded.
        /// </summary>
        public string GetUpdatedPath(string originalPath)
        {
            return pathUpdates.TryGetValue(originalPath, out string updated) ? updated : originalPath;
        }

        /// <summary>
        /// Checks if a decision has already been made for this workbook path.
        /// </summary>
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
        /// <summary>
        /// Finds the highest version of a workbook file.
        /// - If currentPath exists, it finds a strictly NEWER version.
        /// - If currentPath does NOT exist, it finds the highest AVAILABLE version (which may be older).
        /// Returns null if no suitable alternative is found.
        /// </summary>
        public static string FindNewestVersion(string currentPath)
        {
            if (string.IsNullOrEmpty(currentPath))
                return null;

            bool currentFileExists = System.IO.File.Exists(currentPath);
            string normalizedCurrentPath = System.IO.Path.GetFullPath(currentPath);
            string directory = System.IO.Path.GetDirectoryName(normalizedCurrentPath);
            string fileName = System.IO.Path.GetFileNameWithoutExtension(normalizedCurrentPath);
            string extension = System.IO.Path.GetExtension(normalizedCurrentPath);

            var match = System.Text.RegularExpressions.Regex.Match(
                fileName,
                @"^(.+?)[ _]v(\d+)$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (!match.Success) return null; // Not a versioned file

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
                    var versionMatch = System.Text.RegularExpressions.Regex.Match(
                        name,
                        @"[ _]v(\d+)$",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);

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

        /// <summary>
        /// Gets a user-friendly description of the version difference.
        /// </summary>
        public static string GetVersionDescription(string oldPath, string newPath)
        {
            string oldName = System.IO.Path.GetFileName(oldPath);
            string newName = System.IO.Path.GetFileName(newPath);
            return $"{oldName} → {newName}";
        }
    }

    /// <summary>
    /// Caches Excel workbooks during a batch operation to avoid repeated open/close.
    /// Optimized for large workbooks with circular references.
    /// </summary>
    internal class WorkbookCache : IDisposable
    {
        private class WorkbookHandle
        {
            public Excel.Workbook Workbook { get; set; }
            public bool OpenedByUs { get; set; }
            public string ActualPath { get; set; }
            public bool WindowPrepared { get; set; }  // Track if window settings optimized
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
            // A new tracker is created for each batch operation
            this.versionTracker = new WorkbookVersionTracker();
        }

        /// <summary>
        /// Opens a workbook with performance-optimized settings.
        /// Assumes ExcelPerformanceScope is active (AskToUpdateLinks, DisplayAlerts already disabled).
        /// </summary>
        private Excel.Workbook OpenWorkbookOptimized(string path)
        {
            // ExcelPerformanceScope already sets:
            // - AskToUpdateLinks = false
            // - DisplayAlerts = false
            // - Calculation = Manual
            // - ScreenUpdating = false
            // - EnableEvents = false

            try
            {
                // Fast open with all unnecessary features disabled:
                // - UpdateLinks: 0 = Don't update external links (faster, no prompts)
                // - ReadOnly: true = No file lock, faster open
                // - IgnoreReadOnlyRecommended: true = Skip read-only prompt
                // - AddToMru: false = Don't update recent files list
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

                // Fallback to simpler parameters
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

                    // Last resort - minimal open
                    return excelApp.Workbooks.Open(path);
                }
            }
        }

        /// <summary>
        /// Gets a workbook from cache or opens it. Returns null on failure.
        /// This method robustly handles version checks, broken links, and cloud/WOPI paths
        /// (SharePoint, OneDrive, etc.). Also handles cross-user path normalization.
        /// </summary>
        public Excel.Workbook GetOrOpen(string wbPath, out bool newlyOpened, out string actualPath)
        {
            newlyOpened = false;
            actualPath = wbPath;

            if (string.IsNullOrEmpty(wbPath))
                return null;

            // Resolve normalized paths (cross-user cloud sync folder support)
            // %USERPROFILE%\Dropbox\... -> C:\Users\currentuser\Dropbox\...
            if (CloudPathHelper.IsNormalizedPath(wbPath))
            {
                actualPath = CloudPathHelper.ResolvePath(wbPath);
                System.Diagnostics.Debug.WriteLine($"Resolved normalized path: {wbPath} -> {actualPath}");
            }

            bool isCloudPath = CloudPathHelper.IsCloudPath(actualPath);

            // 1. Check if a versioning decision has already been made for this specific path.
            if (versionTracker.WasUpdated(wbPath))
            {
                actualPath = versionTracker.GetUpdatedPath(wbPath);
            }
            // 2. If no decision has been made yet, perform the version check now.
            //    Skip version checking for cloud paths (doesn't apply to SharePoint/OneDrive)
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
                    else // It's an upgrade offer
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
                        // User said No. Record that the original path should be used.
                        versionTracker.RecordUpdate(wbPath, wbPath);
                    }
                }
            }

            // 3. Check the cache for the actual path (or the original normalized path).
            if (cache.TryGetValue(actualPath, out var handle))
            {
                return handle.Workbook;
            }
            // Also check original path if different (for normalized paths)
            if (!string.Equals(wbPath, actualPath, StringComparison.OrdinalIgnoreCase) &&
                cache.TryGetValue(wbPath, out handle))
            {
                return handle.Workbook;
            }

            // 4. Check if the workbook is already open in Excel.
            //    For cloud paths, use flexible matching (URLs can vary slightly)
            Excel.Workbook foundByFilename = null;
            string foundByFilenamePath = null;

            foreach (Excel.Workbook openWb in excelApp.Workbooks)
            {
                try
                {
                    string openWbPath = openWb.FullName;

                    // Use cloud-aware path matching
                    if (CloudPathHelper.PathsMatch(openWbPath, actualPath))
                    {
                        var newHandle = new WorkbookHandle
                        {
                            Workbook = openWb,
                            OpenedByUs = false,
                            ActualPath = openWbPath // Use the actual path from Excel
                        };
                        cache[wbPath] = newHandle;      // Original/normalized path
                        cache[actualPath] = newHandle;  // Resolved path

                        // Also cache under the Excel path if different
                        if (!string.Equals(openWbPath, actualPath, StringComparison.OrdinalIgnoreCase))
                        {
                            cache[openWbPath] = newHandle;
                            System.Diagnostics.Debug.WriteLine($"Cloud workbook found: {actualPath} -> {openWbPath}");
                        }

                        // Update actualPath to what Excel reports
                        actualPath = openWbPath;
                        return openWb;
                    }

                    // Track filename matches as fallback (for cloud/WOPI files)
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

            // 4b. Filename-only fallback for cloud/WOPI files
            //     If path matching failed but we found an open workbook with matching filename,
            //     AND either the stored path OR the open workbook path looks like a cloud path, use it.
            //     This handles cases where paths are stored/reported in different formats.
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
                    cache[wbPath] = newHandle;          // Original/normalized path
                    cache[actualPath] = newHandle;      // Resolved path
                    cache[foundByFilenamePath] = newHandle;

                    // Record path update so future links use the new path
                    versionTracker.RecordUpdate(actualPath, foundByFilenamePath);
                    actualPath = foundByFilenamePath;
                    return foundByFilename;
                }
            }

            // 5. Check if file exists (cloud URLs always return true to attempt open)
            if (!CloudPathHelper.FileExists(actualPath))
            {
                // For cloud URLs that don't exist locally, try to find in sync folders
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

            // 6. Open the file with optimized settings.
            //    For cloud URLs, try multiple approaches:
            //    a) Try opening the URL directly (Excel handles auth if user is signed in)
            //    b) If that fails, try finding and opening the local sync path
            //    c) If all else fails, report error

            Exception lastError = null;
            string[] pathsToTry;

            if (CloudPathHelper.IsCloudUrl(actualPath))
            {
                // For cloud URLs, try: URL first, then search for local sync path
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

                    // Get the actual path Excel reports (may differ for cloud files)
                    string excelReportedPath = wb.FullName;

                    var newHandle = new WorkbookHandle
                    {
                        Workbook = wb,
                        OpenedByUs = true,
                        ActualPath = excelReportedPath
                    };

                    // Cache under multiple keys for flexible lookups:
                    // - Original path (may be normalized %USERPROFILE%\...)
                    // - Resolved/actual path
                    // - Path we tried to open
                    // - Excel's reported path
                    cache[wbPath] = newHandle;
                    cache[actualPath] = newHandle;
                    cache[pathToTry] = newHandle;

                    // Also cache under Excel's reported path if different
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

            // All attempts failed
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

        /// <summary>
        /// Prepares a workbook's window for copy operations by setting optimal display settings.
        /// Call once per workbook - settings are cached to avoid repeated COM calls.
        /// </summary>
        public void PrepareWorkbookForCopy(Excel.Workbook wb)
        {
            if (wb == null) return;

            // Find the handle for this workbook
            WorkbookHandle handle = null;
            foreach (var kvp in cache)
            {
                if (ReferenceEquals(kvp.Value.Workbook, wb))
                {
                    handle = kvp.Value;
                    break;
                }
            }

            // Already prepared?
            if (handle != null && handle.WindowPrepared)
                return;

            // Set window display options for clean copies
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

            // Mark as prepared
            if (handle != null)
                handle.WindowPrepared = true;
        }

        /// <summary>
        /// Overload for backward compatibility.
        /// </summary>
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