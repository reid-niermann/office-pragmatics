using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel.Classes
{
    /// <summary>
    /// Manages the state of a multi-level dependent trace for a worksheet.
    /// This is the "Trace Dependent" version.
    /// </summary>
    internal class TraceStateDependents
    {
        /// <summary>
        /// The address of the range that initiated the trace (e.g., "$A$1" or "$B$2:$C$5").
        /// </summary>
        public string OriginalSelectionAddress { get; set; }

        /// <summary>
        /// The list of dependent cell addresses that were added in the last step.
        /// This is the "frontier" for the next drill-down.
        /// </summary>
        public List<string> LastAddedDependentAddresses { get; set; }

        public TraceStateDependents()
        {
            LastAddedDependentAddresses = new List<string>();
        }
    }

    /// <summary>
    /// Provides functionality to trace dependents with multi-level drill-down for a selected range.
    /// This is the "Trace Dependent" version.
    /// </summary>
    internal static class DependentArrows
    {
        /// <summary>
        /// Stores the tracing state for each worksheet.
        /// Key: Unique sheet identifier (Workbook FullName + "|" + Sheet Name).
        /// Value: The TraceStateDependents object for that sheet.
        /// </summary>
        private static readonly Dictionary<string, TraceStateDependents> _sheetTraceStates = new Dictionary<string, TraceStateDependents>();

        /// <summary>
        /// Drills down one level into the dependents of the selected range.
        /// If this is the first click or the selection has changed, it starts a new trace.
        /// Subsequent clicks on the same selection trace the dependents of the dependents.
        /// </summary>
        public static void DrillDownDependents()
        {
            Excel.Application excelApp = null;
            Excel.Worksheet activeSheet = null;
            Excel.Range selectedRange = null;

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                selectedRange = excelApp.Selection as Excel.Range;
                activeSheet = selectedRange?.Worksheet;

                if (selectedRange == null || activeSheet == null)
                {
                    MessageBox.Show("This command can only be used on a worksheet range.", "Invalid Location", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Unlike precedents, the active cell does not need a formula to have dependents.
                // Any cell can be a precedent, so no formula check is needed here.

                string sheetKey = ((Excel.Workbook)activeSheet.Parent).FullName + "|" + activeSheet.Name;
                _sheetTraceStates.TryGetValue(sheetKey, out TraceStateDependents currentState);

                // --- CONDITION FOR A NEW TRACE ---
                // Start a new trace if no state exists OR if the user has selected a different range.
                if (currentState == null || currentState.OriginalSelectionAddress != selectedRange.Address)
                {
                    StartNewTrace(activeSheet, selectedRange, sheetKey);
                }
                // --- CONDITION FOR DRILL-DOWN ---
                else
                {
                    ContinueTrace(activeSheet, currentState);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An unexpected error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ResetStateForCurrentSheet(excelApp);
            }
            finally
            {
                // Clean up COM objects
                if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);
                if (activeSheet != null) Marshal.ReleaseComObject(activeSheet);
            }
        }

        /// <summary>
        /// Starts a new trace from the selected range (Level 1).
        /// </summary>
        private static void StartNewTrace(Excel.Worksheet activeSheet, Excel.Range selectedRange, string sheetKey)
        {
            activeSheet.ClearArrows();

            var newTraceState = new TraceStateDependents
            {
                OriginalSelectionAddress = selectedRange.Address
            };

            var allDependentAddresses = new HashSet<string>(); // Use HashSet for automatic duplicate removal
            bool screenUpdating = Globals.ThisAddIn.Application.ScreenUpdating;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            try
            {
                // Iterate through each cell in the user's selection
                foreach (Excel.Range cell in selectedRange.Cells)
                {
                    try
                    {
                        // Show the direct dependents. This will throw a COMException if there are none.
                        cell.ShowDependents();

                        // Now, capture the dependents that were just drawn for this cell.
                        Excel.Range dependents = cell.Dependents;
                        foreach (Excel.Range area in dependents.Areas)
                        {
                            allDependentAddresses.Add(area.Address);
                            Marshal.ReleaseComObject(area);
                        }
                        Marshal.ReleaseComObject(dependents);
                    }
                    catch (COMException)
                    {
                        // This specific cell has no dependents on the sheet. This is expected.
                        // We suppress the error and continue to the next cell.
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(cell);
                    }
                }
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = screenUpdating;
            }


            // After checking all cells, if we found nothing, inform the user.
            if (allDependentAddresses.Count == 0)
            {
                MessageBox.Show("The selected cells have no dependents on this sheet.", "No Dependents Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Media.SystemSounds.Beep.Play();
            }

            newTraceState.LastAddedDependentAddresses = allDependentAddresses.ToList();
            _sheetTraceStates[sheetKey] = newTraceState;
        }

        /// <summary>
        /// Continues an existing trace, drilling down one more level.
        /// </summary>
        private static void ContinueTrace(Excel.Worksheet activeSheet, TraceStateDependents currentState)
        {
            if (currentState.LastAddedDependentAddresses.Count == 0)
            {
                // We've already reached the end of the line in the previous step.
                System.Media.SystemSounds.Beep.Play();
                return;
            }

            var newlyFoundAddresses = new HashSet<string>(); // Use HashSet for efficiency and auto-deduplication
            bool screenUpdating = Globals.ThisAddIn.Application.ScreenUpdating;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            try
            {
                // For each of the last-added dependents, show *their* dependents.
                foreach (string address in currentState.LastAddedDependentAddresses)
                {
                    Excel.Range dependentRange = null;
                    try
                    {
                        dependentRange = activeSheet.Range[address];
                        // Call ShowDependents on the *entire* range/area at once.
                        dependentRange.ShowDependents();

                        // Capture the newly added dependents from this range.
                        Excel.Range grandDependents = dependentRange.Dependents;
                        foreach (Excel.Range area in grandDependents.Areas)
                        {
                            newlyFoundAddresses.Add(area.Address);
                            Marshal.ReleaseComObject(area);
                        }
                        Marshal.ReleaseComObject(grandDependents);
                    }
                    catch (COMException)
                    {
                        // This range has no further dependents. Ignore and continue.
                    }
                    finally
                    {
                        if (dependentRange != null) Marshal.ReleaseComObject(dependentRange);
                    }
                }
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = screenUpdating;
            }

            if (newlyFoundAddresses.Count == 0)
            {
                // The drill-down didn't find any new dependents.
                System.Media.SystemSounds.Beep.Play();
            }

            // Update the state's frontier to the new set of dependents.
            currentState.LastAddedDependentAddresses = newlyFoundAddresses.ToList();
        }

        /// <summary>
        /// Removes all precedent and dependent arrows from the active worksheet
        /// and resets the tracing state for dependents.
        /// </summary>
        public static void RemoveAllTracerArrows()
        {
            Excel.Application excelApp = null;
            Excel.Worksheet activeSheet = null;
            try
            {
                excelApp = Globals.ThisAddIn.Application;
                activeSheet = excelApp.ActiveSheet as Excel.Worksheet;
                if (activeSheet == null) return;

                activeSheet.ClearArrows();

                // Also clear our internal state for this sheet.
                ResetStateForCurrentSheet(excelApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while clearing arrows: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (activeSheet != null) Marshal.ReleaseComObject(activeSheet);
            }
        }

        /// <summary>
        /// Helper to reset the state for the current sheet, used for cleanup.
        /// </summary>
        private static void ResetStateForCurrentSheet(Excel.Application excelApp)
        {
            if (excelApp == null) return;
            Excel.Worksheet sheet = null;
            try
            {
                sheet = excelApp.ActiveSheet as Excel.Worksheet;
                if (sheet != null)
                {
                    string key = ((Excel.Workbook)sheet.Parent).FullName + "|" + sheet.Name;
                    if (_sheetTraceStates.ContainsKey(key))
                    {
                        _sheetTraceStates.Remove(key);
                    }
                }
            }
            catch { /* Suppress errors during cleanup */ }
            finally
            {
                if (sheet != null) Marshal.ReleaseComObject(sheet);
            }
        }
    }
}