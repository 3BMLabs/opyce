using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;

namespace powerpoint
{
    public partial class OpycePowerPoint
    {
        opyce.MainRibbon ribbon;
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            opyce.MainRibbon.factory = Globals.Factory.GetRibbonFactory();
            ribbon = new opyce.MainRibbon
            {
                RibbonType = "Microsoft.PowerPoint.Presentation"
            };
            //https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.tools.word.documentbase.createribbonextensibilityobject?view=vsto-2022
            return ribbon.GetExtensibility();
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            opyce.MainRibbon.SetPlaceHolders($"appname=Powerpoint\ninitialization=");
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }
        
        #endregion
    }
}
