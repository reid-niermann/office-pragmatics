using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel
{
    public static class RevertBackwards
    {
        // Static field to store the cell reference before TracePrecedent was called
        private static Excel.Range _savedCell = null;

        /// <summary>
        /// Saves the current active cell reference for later reversion.
        /// This should be called at the start of TracePrecedent.Run() before any navigation occurs.
        /// </summary>
        /// <param name="cell">The cell to save (typically app.ActiveCell)</param>
        public static void SaveCurrentCell(Excel.Range cell)
        {
            // Release the previous saved cell if it exists
            if (_savedCell != null)
            {
                try
                {
                    Marshal.ReleaseComObject(_savedCell);
                }
                catch { }
                _savedCell = null;
            }

            // Save the new cell reference
            if (cell != null)
            {
                _savedCell = cell;
            }
        }

        /// <summary>
        /// Returns to the cell that was saved before TracePrecedent was last called.
        /// This is the main "revert backwards" function.
        /// </summary>
        /// <param name="app">The Excel Application object</param>
        public static void Run(Excel.Application app)
        {
            if (app == null)
            {
                MessageBox.Show("Excel application is not available.", "Revert Backwards",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_savedCell == null)
            {
                MessageBox.Show("No previous cell to revert to. Please run Trace Precedents first.",
                                "Revert Backwards",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Verify the saved cell is still valid
                var testWorksheet = _savedCell.Worksheet;
                var testAddress = _savedCell.Address;

                // Navigate back to the saved cell
                SafeSelect(_savedCell);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to revert to the saved cell. It may have been deleted.\n\nError: {ex.Message}",
                                "Revert Backwards",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);

                // Clear the invalid saved cell
                ClearSavedCell();
            }
        }

        /// <summary>
        /// Clears the saved cell reference. Can be called when you want to reset the history.
        /// </summary>
        public static void ClearSavedCell()
        {
            if (_savedCell != null)
            {
                try
                {
                    Marshal.ReleaseComObject(_savedCell);
                }
                catch { }
                _savedCell = null;
            }
        }

        /// <summary>
        /// Safely selects a range in Excel, handling activation and scrolling.
        /// Same logic as SafeSelect in TracePrecedent.
        /// </summary>
        private static void SafeSelect(Excel.Range rng)
        {
            var app = (Excel.Application)rng.Application;
            bool scrn = app.ScreenUpdating;
            bool evts = app.EnableEvents;

            try
            {
                app.ScreenUpdating = false;
                app.EnableEvents = false;

                Excel.Range target = rng;
                target.Worksheet.Activate();
                app.Goto(target, true);

                try
                {
                    var window = app.ActiveWindow;
                    Excel.Range firstCell = (Excel.Range)target.Cells[1, 1];
                    int targetRow = firstCell.Row;
                    int targetCol = firstCell.Column;
                    int visibleRows = window.VisibleRange.Rows.Count;
                    int visibleCols = window.VisibleRange.Columns.Count;
                    int scrollRow = Math.Max(1, targetRow - visibleRows / 2);
                    int scrollCol = Math.Max(1, targetCol - visibleCols / 2);
                    window.ScrollRow = scrollRow;
                    window.ScrollColumn = scrollCol;
                    Marshal.ReleaseComObject(firstCell);
                }
                catch { }
            }
            finally
            {
                app.EnableEvents = evts;
                app.ScreenUpdating = scrn;
            }
        }
    }
}