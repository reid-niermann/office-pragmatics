using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel.Classes
{
    /// <summary>
    /// Manages the state of a multi-level precedent trace for a worksheet.
    /// </summary>
    internal class TraceState
    {
        /// <summary>
        /// The address of the range that initiated the trace (e.g., "$A$1" or "$B$2:$C$5").
        /// </summary>
        public string OriginalSelectionAddress { get; set; }

        /// <summary>
        /// The list of precedent cell addresses that were added in the last step.
        /// This is the "frontier" for the next drill-down.
        /// </summary>
        public List<string> LastAddedPrecedentAddresses { get; set; }

        public TraceState()
        {
            LastAddedPrecedentAddresses = new List<string>();
        }
    }

    /// <summary>
    /// Provides functionality to trace precedents with multi-level drill-down for a selected range.
    /// </summary>
    internal static class PrecedentArrows
    {
        /// <summary>
        /// Stores the tracing state for each worksheet.
        /// Key: Unique sheet identifier (Workbook FullName + "|" + Sheet Name).
        /// Value: The TraceState object for that sheet.
        /// </summary>
        private static readonly Dictionary<string, TraceState> _sheetTraceStates = new Dictionary<string, TraceState>();

        /// <summary>
        /// Drills down one level into the precedents of the selected range of cells.
        /// If this is the first click or the selection has changed, it starts a new trace.
        /// Subsequent clicks on the same selection trace the precedents of the precedents.
        /// </summary>
        public static void DrillDownPrecedents()
        {
            Excel.Application excelApp = null;
            Excel.Worksheet activeSheet = null;
            Excel.Range selectedRange = null;

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                // Use Selection instead of ActiveCell to support multi-cell ranges
                selectedRange = excelApp.Selection as Excel.Range;
                activeSheet = selectedRange?.Worksheet;

                if (selectedRange == null || activeSheet == null)
                {
                    MessageBox.Show("This command can only be used on a worksheet range.", "Invalid Location", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Check if any cell in the selection has a formula.
                // The HasFormula property returns 'false' if no cells have formulas, 'true' if all do, and 'null' for a mix.
                // We must use .Equals() for safe comparison because HasFormula returns an object.
                if (false.Equals(selectedRange.HasFormula))
                {
                    MessageBox.Show("None of the selected cells contain a formula.", "No Formulas Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    System.Media.SystemSounds.Exclamation.Play();
                    return;
                }

                string sheetKey = ((Excel.Workbook)activeSheet.Parent).FullName + "|" + activeSheet.Name;
                _sheetTraceStates.TryGetValue(sheetKey, out TraceState currentState);

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

            var newTraceState = new TraceState
            {
                OriginalSelectionAddress = selectedRange.Address
            };

            var allPrecedentAddresses = new HashSet<string>(); // Use HashSet for automatic duplicate removal
            bool screenUpdating = Globals.ThisAddIn.Application.ScreenUpdating;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            try
            {
                // Iterate through each cell in the user's selection
                foreach (Excel.Range cell in selectedRange.Cells)
                {
                    try
                    {
                        // Show direct precedents only if the cell has a formula
                        if (true.Equals(cell.HasFormula)) // Safe comparison for single cell
                        {
                            cell.ShowPrecedents();

                            // Capture the precedents that were just drawn for this cell
                            Excel.Range precedents = cell.Precedents;
                            foreach (Excel.Range area in precedents.Areas)
                            {
                                allPrecedentAddresses.Add(area.Address);
                                Marshal.ReleaseComObject(area);
                            }
                            Marshal.ReleaseComObject(precedents);
                        }
                    }
                    catch (COMException)
                    {
                        // This means the formula has no range precedents (e.g., =TODAY()) or other error. Ignore and continue.
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


            if (allPrecedentAddresses.Count == 0)
            {
                System.Media.SystemSounds.Beep.Play(); // Beep if no precedents were found at all
            }

            newTraceState.LastAddedPrecedentAddresses = allPrecedentAddresses.ToList();
            _sheetTraceStates[sheetKey] = newTraceState;
        }

        /// <summary>
        /// Continues an existing trace, drilling down one more level.
        /// </summary>
        private static void ContinueTrace(Excel.Worksheet activeSheet, TraceState currentState)
        {
            if (currentState.LastAddedPrecedentAddresses.Count == 0)
            {
                // We've already reached the end of the line on the previous step.
                System.Media.SystemSounds.Beep.Play();
                return;
            }

            var newlyFoundAddresses = new HashSet<string>(); // Use HashSet for efficiency
            bool screenUpdating = Globals.ThisAddIn.Application.ScreenUpdating;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            try
            {
                // For each of the last-added precedents, show *their* precedents.
                foreach (string address in currentState.LastAddedPrecedentAddresses)
                {
                    Excel.Range precedentRange = null;
                    try
                    {
                        precedentRange = activeSheet.Range[address];
                        // Call ShowPrecedents on the *entire* range/area at once.
                        precedentRange.ShowPrecedents();

                        // Capture the newly added precedents from this range.
                        Excel.Range grandPrecedents = precedentRange.Precedents;
                        foreach (Excel.Range area in grandPrecedents.Areas)
                        {
                            newlyFoundAddresses.Add(area.Address);
                            Marshal.ReleaseComObject(area);
                        }
                        Marshal.ReleaseComObject(grandPrecedents);
                    }
                    catch (COMException)
                    {
                        // This range has no further precedents. Ignore and continue.
                    }
                    finally
                    {
                        if (precedentRange != null) Marshal.ReleaseComObject(precedentRange);
                    }
                }
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = screenUpdating;
            }

            if (newlyFoundAddresses.Count == 0)
            {
                // The drill-down didn't find any new precedents across all frontiers.
                System.Media.SystemSounds.Beep.Play();
            }

            // Update the state's frontier to the new set of precedents.
            currentState.LastAddedPrecedentAddresses = newlyFoundAddresses.ToList();
        }

        /// <summary>
        /// Removes all precedent and dependent arrows from the active worksheet
        /// and resets the tracing state.
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