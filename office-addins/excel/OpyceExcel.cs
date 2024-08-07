using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using opyce;
using System.IO;

namespace excel
{
    public partial class OpyceExcel
    {
        opyce.MainRibbon ribbon;
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            opyce.MainRibbon.factory = Globals.Factory.GetRibbonFactory();
            ribbon = new opyce.MainRibbon();
            return ribbon.GetExtensibility();
        }
        void Application_WorkbookOpen(Excel.Workbook wb)
        {
            var customData = Serializer.GetCustomXmlPart<CustomXML>(wb, Serializer.OpyceNameSpace);
            if (customData != null)
            {
                customData.Value = "blah";
            }
            opyce.MainRibbon.SetPlaceHolders($"appname=Excel\ninitialization=self.workbook = self.app.Workbooks(\"{this.Application.ActiveWorkbook.Name}\")");
        }
        void Application_WorkbookBeforeSave(Excel.Workbook wb, bool SaveAsUI, ref bool Cancel)
        {
            //if(File.Exists(MainRibbon.))
            //string mainFile = 
            //var customData = new CustomXML { Key = "mainFile", Value = "exampleValue" };
            //Serializer.AddCustomXmlPart(wb, customData, Serializer.OpyceNameSpace);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            //this.Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
            this.Application.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(Application_WorkbookOpen);
            this.Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
        }

        #endregion
    }
}
