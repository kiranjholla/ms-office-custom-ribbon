using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace Excel_Custom_Ribbon
{
    [ComVisible(true)]
    public class ExcelRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public ExcelRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Excel_Custom_Ribbon.ExcelRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void LoadExcelRibbon(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        public void setDateFormat(Office.IRibbonControl control)
        {
            string style = "Date";
            string format = "[$-409]d-mmm-yyyy_);@_)";
            applyStyle(style, format);
        }

        public void setRsFormat(Office.IRibbonControl control)
        {
            string style = "Currency Rs.";
            string format = "_(\"Rs.\"* #,##0.00_);_(\"Rs.\"* - #,##0.00_);_(\"Rs.\"* - ??_);_(@_)";
            applyStyle(style, format);
        }
        
        public void setINRFormat(Office.IRibbonControl control)
        {
            string style = "Currency ₹";
            string format = "_(₹* #,##0.00_);_(₹* - #,##0.00_);_(₹* - ??_);_(@_)";
            applyStyle(style, format);
        }

        public void setPoundFormat(Office.IRibbonControl control)
        {
            string style = "Currency £";
            string format = "_(£* #,##0.00_);_(£* - #,##0.00_);_(£* - ??_);_(@_)";
            applyStyle(style, format);
        }
        public void setEuroFormat(Office.IRibbonControl control)
        {
            string style = "Currency €";
            string format = "_(€* #,##0.00_);_(€* - #,##0.00_);_(€* - ??_);_(@_)";
            applyStyle(style, format);
        }

        public stdole.IPictureDisp GetImage(string imageName)
        {
            switch (imageName)
            {
                case "Euro":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.Euro);

                case "Pound":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.Pound);

                case "Rs":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.Rs);

                case "Inr":
                    return PictureConverter.ImageToPictureDisp(Properties.Resources.Inr);

                default:
                    return null;
            }
        }

        #endregion

        #region Helpers

        private void applyStyle(string style, string format)
        {
            Excel.Range xRange = Globals.ExcelCustomRibbon.Application.Selection;
            try
            {
                xRange.Style = style;

            }
            catch (Exception e) when (e.Message.Contains("Style '" + style + "' not found."))
            {
                Excel.Style dateStyle = Globals.ExcelCustomRibbon.Application.ActiveWorkbook.Styles.Add(style);
                dateStyle.IncludeAlignment = false;
                dateStyle.IncludeBorder = false;
                dateStyle.IncludeFont = false;
                dateStyle.IncludeNumber = true;
                dateStyle.NumberFormat = format;
                dateStyle.IncludePatterns = false;
                dateStyle.IncludeProtection = false;

                xRange.Style = style;
            }
        }

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
