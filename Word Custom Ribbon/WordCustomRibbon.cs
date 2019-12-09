using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace Word_Custom_Ribbon
{
    public partial class WordCustomRibbon
    {
        private void WordCustomRibbon_Startup(object sender, System.EventArgs e)
        {
        }

        private void WordCustomRibbon_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new WordRibbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(WordCustomRibbon_Startup);
            this.Shutdown += new System.EventHandler(WordCustomRibbon_Shutdown);
        }
        
        #endregion
    }
}
