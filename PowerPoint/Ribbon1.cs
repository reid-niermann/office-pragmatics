// PowerPoint\Ribbon1.cs
using Microsoft.Office.Core;
using OfficePragmatics_Powerpoint.Classes.ExcelLinking;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms; // Required for MessageBox
using Office = Microsoft.Office.Core;

namespace OfficePragmatics_Powerpoint
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private readonly RibbonController controller;
        private int _linkMatchSelectedIndex;
        public static Ribbon1 Instance { get; private set; }

        public Ribbon1()
        {
            this.controller = new RibbonController();
            _linkMatchSelectedIndex = 2; // Default to "Match All"
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OfficePragmatics_Powerpoint.Ribbon1.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            Instance = this;
        }

        public static void TryInvalidate()
        {
            try
            {
                Instance?.ribbon?.Invalidate();
            }
            catch { }
        }

        #region Callbacks Delegated to Controller

        public Bitmap GetCustomImage(IRibbonControl control) => controller.GetCustomImage(control);
        public string GetScreentip(IRibbonControl control) => controller.GetScreentip(control);
        public string GetSupertip(IRibbonControl control) => controller.GetSupertip(control);
        public string GetKeytip(IRibbonControl control) => controller.GetKeytip(control);
        public void OnUniversalAction(IRibbonControl control) => controller.OnUniversalAction(control);
        public string GetLabel(IRibbonControl control) => controller.GetLabel(control);

        #endregion

        #region Callbacks for galLinkMatch

        public int GetLinkItemCount(Office.IRibbonControl _) => 3;

        public Bitmap GetLinkGalleryImage(Office.IRibbonControl _)
        {
            switch (_linkMatchSelectedIndex)
            {
                case 0: return Properties.Resources.btnSameWidth;
                case 1: return Properties.Resources.btnSameHeight;
                case 2: return Properties.Resources.btnSameHW;
                default: return Properties.Resources.TBU;
            }
        }

        // Corrected: The 'selectedId' parameter was unused and is now renamed to '_'
        public void OnLinkGalleryItemClicked(Office.IRibbonControl control, string _, int selectedIndex)
        {
            _linkMatchSelectedIndex = selectedIndex;

            LinkMatchMode mode;
            switch (selectedIndex)
            {
                case 0: mode = LinkMatchMode.WidthOnly; break;
                case 1: mode = LinkMatchMode.HeightOnly; break;
                default: mode = LinkMatchMode.Both; break;
            }
            LinkMatcher.MatchLinks(mode);

            ribbon?.InvalidateControl(control.Id);
        }

        public int GetLinkSelectedItemIndex(Office.IRibbonControl _)
        {
            return _linkMatchSelectedIndex;
        }

        public string GetLinkMatchItemLabel(Office.IRibbonControl _, int index)
        {
            string check = (index == _linkMatchSelectedIndex) ? "✓ " : "";
            switch (index)
            {
                case 0: return check + "Same Width";
                case 1: return check + "Same Height";
                case 2: return check + "Match All";
                default: return string.Empty;
            }
        }

        // Corrected: Restored the original method signature that other classes depend on.
        public string GetCurrentImportMatchType()
        {
            switch (_linkMatchSelectedIndex)
            {
                case 0: return "Match Width";
                case 1: return "Match Height";
                case 2: return "Match All";
                default: return "Match All"; // Fallback to a sensible default
            }
        }

        #endregion

        #region Helper to load XML

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            foreach (string name in resourceNames)
            {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(name)))
                    {
                        return resourceReader.ReadToEnd();
                    }
                }
            }
            return null;
        }

        #endregion
    }

    public enum LinkMatchMode { WidthOnly, HeightOnly, Both }

    public static class LinkMatcher
    {
        public static void MatchLinks(LinkMatchMode mode)
        {
            //MessageBox.Show($"Matching links with mode: {mode}");
        }
    }
}