using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;

namespace opyce
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        public static RibbonFactory factory = null;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
            : base(factory ?? Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        public IRibbonExtensibility GetExtensibility()
        {
            return factory.CreateRibbonManager(
                new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { this });
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.openInPython = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Opyce";
            this.tab1.Name = "tab1";
            this.tab1.Position = this.Factory.RibbonPosition.AfterOfficeId("TabHome");
            // 
            // group1
            // 
            this.group1.Items.Add(this.openInPython);
            this.group1.Label = "opyce";
            this.group1.Name = "group1";
            // 
            // openInPython
            // 
            this.openInPython.Label = "open in python";
            this.openInPython.Name = "openInPython";
            this.openInPython.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenInPythonButton_Click);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal RibbonButton openInPython;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
