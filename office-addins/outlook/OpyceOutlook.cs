﻿using Microsoft.Office.Core;

namespace outlook
{
    public partial class OpyceOutlook
    {
        opyce.MainRibbon ribbon;
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            opyce.MainRibbon.factory = Globals.Factory.GetRibbonFactory();
            ribbon = new opyce.MainRibbon();
            return ribbon.GetExtensibility(); // Assuming Ribbon1 is the class name in the second VSTO add-in
        }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            opyce.MainRibbon.SetPlaceHolders("Outlook");
            ribbon.Serialize(false);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += OpyceOutlook_Shutdown;
        }

        private void OpyceOutlook_Shutdown(object sender, System.EventArgs e)
        {
        }

        #endregion
    }
}
