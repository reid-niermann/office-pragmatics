//  XlamInstaller.cs
//  OfficePragmatics – Excel VSTO add-in
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel.Helpers
{
    /// <summary>
    /// Manages the deployment and updating of the companion VBA XLAM.
    /// Handles file locking by unloading the add-in before attempting updates.
    /// </summary>
    internal static class XlamInstaller
    {
        private const string XlamFileName = "Pragmatics Companion.xlam";
        // Must match the embedded resource name exactly
        private const string ResName = "OfficePragmatics_Excel.CompanionXLAM." + XlamFileName;

        private static string _targetPath;
        private static byte[] _embeddedBytes;
        private static bool _updateRequired;

        /// <summary>
        /// Call this from ThisAddIn_Startup.
        /// It prepares the installer but defers COM operations until Excel is ready.
        /// </summary>
        internal static void EnsureInstalled(Excel.Application app)
        {
            try
            {
                // 1. Setup Paths
                string addInsPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Microsoft", "AddIns");

                Directory.CreateDirectory(addInsPath);
                _targetPath = Path.Combine(addInsPath, XlamFileName);

                // 2. Read Resource
                _embeddedBytes = ReadEmbedded();
                if (_embeddedBytes == null) return; // Resource missing

                // 3. Check if update is needed (Hash Check)
                // We do NOT write yet if the file exists, because Excel might have it locked.
                if (!File.Exists(_targetPath))
                {
                    // Safe to write immediately if file doesn't exist
                    File.WriteAllBytes(_targetPath, _embeddedBytes);
                    _updateRequired = false;
                }
                else
                {
                    // File exists. Compare hashes to see if we need to overwrite.
                    bool isDifferent = !BytesEqual(_embeddedBytes, File.ReadAllBytes(_targetPath));
                    _updateRequired = isDifferent;
                }

                // 4. Hook into Excel to perform the actual Load/Update
                // We use WorkbookActivate because it fires when the UI is ready and COM is stable.
                app.WorkbookActivate += Application_WorkbookActivate;
            }
            catch (Exception)
            {
                // Log error silently or fallback
            }
        }

        /// <summary>
        /// Fires once when the first workbook is active. 
        /// This is the safe time to manipulate AddIns and write files.
        /// </summary>
        private static void Application_WorkbookActivate(Excel.Workbook wb)
        {
            Excel.Application app = wb.Application;

            // Unhook immediately so this only runs once per session
            app.WorkbookActivate -= Application_WorkbookActivate;

            try
            {
                Excel.AddIn addIn = FindAddIn(app, XlamFileName);

                if (_updateRequired)
                {
                    // --- UPDATE SCENARIO ---

                    // 1. Release the Lock: If it's currently installed, unload it.
                    if (addIn != null && addIn.Installed)
                    {
                        addIn.Installed = false;
                    }

                    // 2. Overwrite the file (Lock should be released now)
                    try
                    {
                        File.WriteAllBytes(_targetPath, _embeddedBytes);
                    }
                    catch (IOException)
                    {
                        // If we still can't write (permission/AV), we abort the update 
                        // but proceed to ensure the old version is at least loaded.
                    }

                    // 3. Refetch the object (sometimes necessary after file swap) or Add if missing
                    if (addIn == null)
                    {
                        addIn = app.AddIns.Add(_targetPath, false);
                    }

                    // 4. Reload
                    addIn.Installed = true;
                }
                else
                {
                    // --- NORMAL STARTUP SCENARIO ---
                    // Just ensure it's loaded
                    if (addIn == null)
                    {
                        addIn = app.AddIns.Add(_targetPath, false);
                    }

                    if (!addIn.Installed)
                    {
                        addIn.Installed = true;
                    }
                }
            }
            catch (Exception)
            {
                // Swallow errors here to prevent crashing Excel startup
            }
            finally
            {
                // Clear memory
                _embeddedBytes = null;
            }
        }

        /*─────────────────────── Helpers ────────────────────────────────*/

        private static Excel.AddIn FindAddIn(Excel.Application app, string fileName)
        {
            // Note: We search by Name (filename), not FullName, to be more robust 
            // if the user moved their AppData folder, though standard installs usually don't.
            foreach (Excel.AddIn a in app.AddIns)
            {
                if (string.Equals(a.Name, fileName, StringComparison.OrdinalIgnoreCase))
                {
                    return a;
                }
            }
            return null;
        }

        private static byte[] ReadEmbedded()
        {
            using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResName))
            {
                if (s == null) return null;
                using (var ms = new MemoryStream())
                {
                    s.CopyTo(ms);
                    return ms.ToArray();
                }
            }
        }

        private static bool BytesEqual(byte[] a, byte[] b)
        {
            if (a.Length != b.Length) return false;
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hashA = sha.ComputeHash(a);
                byte[] hashB = sha.ComputeHash(b);
                return hashA.SequenceEqual(hashB);
            }
        }
    }
}