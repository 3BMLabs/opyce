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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.openInPython = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.pyfunc1 = this.Factory.CreateRibbonButton();
            this.pyfunc2 = this.Factory.CreateRibbonButton();
            this.pyfunc3 = this.Factory.CreateRibbonButton();
            this.pyfunc4 = this.Factory.CreateRibbonButton();
            this.pyfunc5 = this.Factory.CreateRibbonButton();
            this.pyfunc6 = this.Factory.CreateRibbonButton();
            this.pyfunc7 = this.Factory.CreateRibbonButton();
            this.pyfunc8 = this.Factory.CreateRibbonButton();
            this.pyfunc9 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
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
            // group2
            // 
            this.group2.Items.Add(this.pyfunc1);
            this.group2.Items.Add(this.pyfunc2);
            this.group2.Items.Add(this.pyfunc3);
            this.group2.Items.Add(this.pyfunc4);
            this.group2.Items.Add(this.pyfunc5);
            this.group2.Items.Add(this.pyfunc6);
            this.group2.Items.Add(this.pyfunc7);
            this.group2.Items.Add(this.pyfunc8);
            this.group2.Items.Add(this.pyfunc9);
            this.group2.Label = "functions";
            this.group2.Name = "group2";
            // 
            // pyfunc1
            // 
            this.pyfunc1.Label = "pyfunc1";
            this.pyfunc1.Name = "pyfunc1";
            this.pyfunc1.Visible = false;
            this.pyfunc1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.executePythonFunction);
            // 
            // pyfunc2
            // 
            this.pyfunc2.Label = "pyfunc2";
            this.pyfunc2.Name = "pyfunc2";
            this.pyfunc2.Visible = false;
            this.pyfunc2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.executePythonFunction);
            // 
            // pyfunc3
            // 
            this.pyfunc3.Label = "pyfunc3";
            this.pyfunc3.Name = "pyfunc3";
            this.pyfunc3.Visible = false;
            this.pyfunc3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.executePythonFunction);
            // 
            // pyfunc4
            // 
            this.pyfunc4.Label = "pyfunc4";
            this.pyfunc4.Name = "pyfunc4";
            this.pyfunc4.Visible = false;
            this.pyfunc4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.executePythonFunction);
            // 
            // pyfunc5
            // 
            this.pyfunc5.Label = "pyfunc5";
            this.pyfunc5.Name = "pyfunc5";
            this.pyfunc5.Visible = false;
            this.pyfunc5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.executePythonFunction);
            // 
            // pyfunc6
            // 
            this.pyfunc6.Label = "pyfunc6";
            this.pyfunc6.Name = "pyfunc6";
            this.pyfunc6.Visible = false;
            this.pyfunc6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.executePythonFunction);
            // 
            // pyfunc7
            // 
            this.pyfunc7.Label = "pyfunc7";
            this.pyfunc7.Name = "pyfunc7";
            this.pyfunc7.Visible = false;
            this.pyfunc7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.executePythonFunction);
            // 
            // pyfunc8
            // 
            this.pyfunc8.Label = "pyfunc8";
            this.pyfunc8.Name = "pyfunc8";
            this.pyfunc8.Visible = false;
            this.pyfunc8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.executePythonFunction);
            // 
            // pyfunc9
            // 
            this.pyfunc9.Label = "pyfunc9";
            this.pyfunc9.Name = "pyfunc9";
            this.pyfunc9.Visible = false;
            this.pyfunc9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.executePythonFunction);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = resources.GetString("$this.RibbonType");
            this.Tabs.Add(this.tab1);
            this.Close += new System.EventHandler(this.MainRibbon_Close);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal RibbonButton openInPython;
        internal RibbonGroup group2;
        internal RibbonButton pyfunc1;
        internal RibbonButton pyfunc2;
        internal RibbonButton pyfunc3;
        internal RibbonButton pyfunc4;
        internal RibbonButton pyfunc5;
        internal RibbonButton pyfunc6;
        internal RibbonButton pyfunc7;
        internal RibbonButton pyfunc8;
        internal RibbonButton pyfunc9;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
