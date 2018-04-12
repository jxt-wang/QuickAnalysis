using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace QuickAnalysis
{
    public partial class ThisAddIn
    {

        private ImportData_Settings settings_importData;
        private Microsoft.Office.Tools.CustomTaskPane pane_importData;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            settings_importData = new ImportData_Settings();
            pane_importData = this.CustomTaskPanes.Add(settings_importData, "Import Settings");
            settings_importData.Visible = true;
            //settings_importData.Visible = true; //TODO: ADJUST FOR EVENT 
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new QuickAnalysis_UI();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
