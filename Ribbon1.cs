using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioCustomMenu
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        /// <summary>
        /// The ribbon (top menu) handler, just drops the instructions to the page and the test shapes
        /// </summary>
        /// <param name="control">the ribbon button</param>
        public void OnClickMe(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;
            app.Documents.Add("");
            var stn = app.Documents.OpenEx("basic.vss", (short)Visio.VisOpenSaveArgs.visOpenDocked + (short)Visio.VisOpenSaveArgs.visOpenRO);
            
            var rect = app.ActivePage.DrawRectangle(0, 5, 5, 7);
            rect.Text = "Try using context menu on the shapes";

            app.ActivePage.Drop(stn.Masters.ItemU["Rectangle"], 2, 2);
            app.ActivePage.Drop(stn.Masters.ItemU["Circle"], 4, 2);
        }

        /// <summary>
        /// menu handler that is called when user click context menu item
        /// </summary>
        /// <param name="control">the context menu item clicked</param>
        public void OnContextMenuItemClicked(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;

            var scopeId = app.BeginUndoScope("Paint Shape Red and Put Text");
            try
            {
                var sel = app.ActiveWindow.Selection;
                foreach (Visio.Shape shp in sel)
                {
                    shp.Text = "Hello";
                    shp.Cells["FillForegnd"].Formula = "2";
                }
                app.EndUndoScope(scopeId, true);
            }
            catch (Exception e)
            {
                app.EndUndoScope(scopeId, false);
                System.Diagnostics.Trace.TraceError(e.ToString());
            }
        }

        /// <summary>
        /// callback to return if the menu item should be enabled or disabled
        /// for demo purposes, it enables "Rect" menu item if any rect is selected and "circle" if any circle
        /// it does so by examinint "tag" property of the menu item (set in the XML)
        /// </summary>
        /// <param name="control">the menu item</param>
        /// <returns></returns>
        public bool OnContextMenuItemEnabled(Office.IRibbonControl control)
        {
            return control.Tag == "RECT"  ? HaveSelected("Rectangle") : HaveSelected("Circle");
        }

        /// <summary>
        /// callback to return if the menu item should be visible or hidden
        /// for demo purposes, it show the custom menu items if either "circle" or "rectangle" shape is selected
        /// </summary>
        /// <param name="control">the menu item</param>
        /// <returns></returns>
        public bool OnContextMenuItemVisible(Office.IRibbonControl control)
        {
            return HaveSelected("Rectangle") || HaveSelected("Circle");
        }

        bool HaveSelected(string name)
        {
            var app = Globals.ThisAddIn.Application;
            return app.ActiveWindow.Selection.Cast<Visio.Shape>().Any(s => s?.Master?.Name == name);
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("VisioCustomMenu.Ribbon1.xml");
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
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
