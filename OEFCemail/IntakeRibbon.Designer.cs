﻿
namespace OEFCemail
{
    partial class IntakeRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public IntakeRibbon()
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
            Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
            this.toggleButtonIntakeDisplay = this.Factory.CreateRibbonToggleButton();
            this.Tab1 = this.Factory.CreateRibbonTab();
            group1 = this.Factory.CreateRibbonGroup();
            group1.SuspendLayout();
            this.Tab1.SuspendLayout();
            this.SuspendLayout();
            // 
            // group1
            // 
            group1.Items.Add(this.toggleButtonIntakeDisplay);
            group1.Label = "Show Intake Panel";
            group1.Name = "group1";
            // 
            // toggleButtonIntakeDisplay
            // 
            this.toggleButtonIntakeDisplay.Label = "Toggel Intake Panel";
            this.toggleButtonIntakeDisplay.Name = "toggleButtonIntakeDisplay";
            this.toggleButtonIntakeDisplay.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonIntakeDisplay_Click);
            // 
            // Tab1
            // 
            this.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Tab1.Groups.Add(group1);
            this.Tab1.Label = "Email Intake";
            this.Tab1.Name = "Tab1";
            // 
            // IntakeRibbon
            // 
            this.Name = "IntakeRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.Tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.IntakeRibbon_Load);
            group1.ResumeLayout(false);
            group1.PerformLayout();
            this.Tab1.ResumeLayout(false);
            this.Tab1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonIntakeDisplay;
    }

    partial class ThisRibbonCollection
    {
        internal IntakeRibbon IntakeRibbon
        {
            get { return this.GetRibbon<IntakeRibbon>(); }
        }
    }
}
