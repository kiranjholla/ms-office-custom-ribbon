using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace Excel_Custom_Ribbon
{
    public partial class ExcelCustomRibbon
    {
        private void ExcelCustomRibbon_Startup(object sender, System.EventArgs e)
        {
        }

        private void ExcelCustomRibbon_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new ExcelRibbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ExcelCustomRibbon_Startup);
            this.Shutdown += new System.EventHandler(ExcelCustomRibbon_Shutdown);
        }
        
        #endregion
    }
}
