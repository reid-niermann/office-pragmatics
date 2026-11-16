using System;
using Excel = Microsoft.Office.Interop.Excel;
using OfficePragmatics_Excel.Settings;

namespace OfficePragmatics_Excel
{
    public partial class ThisAddIn
    {
        private AddInUtilities _utilities;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        protected override object RequestComAddInAutomationService()
        {
            // Ensure utilities exists if external COM callers (like VBA) request it
            if (_utilities == null)
                _utilities = new AddInUtilities();
            return _utilities;
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            if (_utilities == null) _utilities = new AddInUtilities();

            // 1. Hook Events FIRST
            // These ensure that as soon as Excel is actually ready (Book1 opens), we register keys.
            this.Application.WorkbookOpen += Application_WorkbookOpen;
            this.Application.WorkbookActivate += Application_WorkbookActivate;

            // Re-hook sheet change for auto-color features
            UpdateSheetChangeHandler();

            try
            {
                // 2. Companion XLAM install check (Non-UI, safe to run anytime)
                try { OfficePragmatics_Excel.Helpers.XlamInstaller.EnsureInstalled(this.Application); } catch { }

                // 3. Initial Registration Attempt
                // If Excel started with a file (double-clicked a file), this might work immediately.
                // If Excel started blank (Book1), this might fail silently (Count=0),
                // but the WorkbookOpen event below will catch it milliseconds later.
                _utilities.RegisterShortcuts();

                // 4. Initialize Settings
                SettingsRuntime.InitializeOnAddInStartup();

                // 5. Safe handshake with XLAM (Push connection)
                // Note: With the new "Lazy Load" VBA code, this is less critical but still good practice.
                if (IsXlamLoaded())
                {
                    try { this.Application.Run("InitPragmaticsAddin"); } catch { }
                }
            }
            catch (Exception)
            {
                // Log error if you have logging, otherwise fail silently to allow Excel to load
            }
        }

        private void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            // Standard Event: Just ensure shortcuts are active when a book opens
            RegisterShortcutsSafe();
        }

        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            // Standard Event: Ensure shortcuts persist when switching books
            RegisterShortcutsSafe();
            Ribbon1.TryInvalidate();
        }

        private void RegisterShortcutsSafe()
        {
            try
            {
                if (_utilities == null) _utilities = new AddInUtilities();

                // This calls your RegisterShortcuts logic.
                // Ensure your AddInUtilities.RegisterShortcuts() still has the check:
                // "if (app.Workbooks.Count == 0) return;" 
                // This prevents errors if this event fires weirdly early.
                _utilities.RegisterShortcuts();
            }
            catch { }
        }

        private void Application_SheetChange(object Sh, Excel.Range Target)
        {
            // Guard clause to prevent recursive loops
            Application.EnableEvents = false;
            try
            {
                FontColor.AutoColorRange(Target);
            }
            finally
            {
                Application.EnableEvents = true;
            }
        }

        public void UpdateSheetChangeHandler()
        {
            // Unsubscribe first to ensure we don't subscribe twice
            Application.SheetChange -= Application_SheetChange;

            if (SettingsForm.CurrentSettings.AutoFontColorEnabled)
            {
                Application.SheetChange += Application_SheetChange;
            }
        }

        private bool IsXlamLoaded()
        {
            try
            {
                foreach (Excel.AddIn addIn in this.Application.AddIns)
                {
                    if (string.Equals(addIn.Name, "Pragmatics Companion.xlam", StringComparison.OrdinalIgnoreCase))
                    {
                        return addIn.Installed;
                    }
                }
            }
            catch { }
            return false;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Cleanup if necessary
            _utilities = null;
        }

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}