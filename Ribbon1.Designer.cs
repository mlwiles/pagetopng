﻿namespace PageToPNG
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
            this.PTP = this.Factory.CreateRibbonGroup();
            this.PageToPNGBTN = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.PTP.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.PTP);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // PTP
            // 
            this.PTP.Items.Add(this.PageToPNGBTN);
            this.PTP.Label = "PTP";
            this.PTP.Name = "PTP";
            // 
            // PageToPNGBTN
            // 
            this.PageToPNGBTN.Label = "PageToPNG";
            this.PageToPNGBTN.Name = "PageToPNGBTN";
            this.PageToPNGBTN.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PageToPNGBTN_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.PTP.ResumeLayout(false);
            this.PTP.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup PTP;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PageToPNGBTN;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
