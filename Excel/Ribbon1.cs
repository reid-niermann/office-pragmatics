using Microsoft.Office.Core;
using OfficePragmatics_Excel.Features;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace OfficePragmatics_Excel
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private readonly RibbonController controller;

        public Ribbon1()
        {
            this.controller = new RibbonController();
        }

        private static Ribbon1 _lastLoaded;

        public static void TryInvalidate()
        {
            try { _lastLoaded?.ribbon?.Invalidate(); } catch { }
        }

        #region IRibbonExtensibility Members
        public string GetCustomUI(string ribbonID) => GetResourceText("OfficePragmatics_Excel.Ribbon1.xml");
        #endregion

        #region Ribbon Callbacks (Delegated)
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            _lastLoaded = this;
        }

        public Bitmap GetCustomImage(IRibbonControl control) => controller.GetCustomImage(control);
        public string GetScreentip(IRibbonControl control) => controller.GetScreentip(control);
        public string GetSupertip(IRibbonControl control) => controller.GetSupertip(control);
        public string GetKeytip(IRibbonControl control) => controller.GetKeytip(control);
        public void OnUniversalAction(IRibbonControl control) => controller.OnUniversalAction(control);
        public string GetLabel(IRibbonControl control) => controller.GetLabel(control);


        public void OnInsertSymbol(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            SymbolGallery.InsertByIndex(selectedIndex);
        }
        #endregion

        #region Helpers
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
}