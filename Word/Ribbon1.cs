// Ribbon1.cs
using Microsoft.Office.Core;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace Office_Pragmatics_Word
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private readonly RibbonController controller;
        public static Ribbon1 Instance { get; private set; }

        public Ribbon1()
        {
            this.controller = new RibbonController();
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Office_Pragmatics_Word.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks

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


        #endregion

        #region Callbacks Delegated to Controller

        public Bitmap GetCustomImage(IRibbonControl control) => controller.GetCustomImage(control);
        public string GetScreentip(IRibbonControl control) => controller.GetScreentip(control);
        public string GetSupertip(IRibbonControl control) => controller.GetSupertip(control);
        public string GetKeytip(IRibbonControl control) => controller.GetKeytip(control);
        public void OnUniversalAction(IRibbonControl control) => controller.OnUniversalAction(control);
        public string GetLabel(IRibbonControl control) => controller.GetLabel(control);

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
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}