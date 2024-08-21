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
        void onOpen(Excel.Workbook wb)
        {
            ribbon.Serialize(false, wb);
            opyce.MainRibbon.SetPlaceHolders($"appname=Excel\ninitialization=self.workbook = self.app.Workbooks(\"{this.Application.ActiveWorkbook.Name}\")");
        }
        void OnSave(Excel.Workbook wb, bool SaveAsUI, ref bool Cancel)
        {
            ribbon.Serialize(true, wb);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Application.WorkbookActivate += onOpen;
            this.Application.WorkbookBeforeSave += OnSave;
        }


        #endregion
    }
}
