﻿namespace ExcelAddIn2
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
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
            this.Export = this.Factory.CreateRibbonButton();
            this.AutoFit = this.Factory.CreateRibbonButton();
            this.Clear = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.Export);
            this.group1.Items.Add(this.AutoFit);
            this.group1.Items.Add(this.Clear);
            this.group1.Label = "Awesome Buttons";
            this.group1.Name = "group1";
            // 
            // Export
            // 
            this.Export.Label = "Export to CSV";
            this.Export.Name = "Export";
            this.Export.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.exportCsv_Click);
            // 
            // AutoFit
            // 
            this.AutoFit.Label = "AutoFit";
            this.AutoFit.Name = "AutoFit";
            this.AutoFit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AutoFit_Click);
            // 
            // Clear
            // 
            this.Clear.Label = "Clear ALL";
            this.Clear.Name = "Clear";
            this.Clear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Clear_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Export;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AutoFit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Clear;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
