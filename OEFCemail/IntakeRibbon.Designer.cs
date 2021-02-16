
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
            Microsoft.Office.Tools.Ribbon.RibbonGroup Test;
            this.IntakeTab = this.Factory.CreateRibbonTab();
            Test = this.Factory.CreateRibbonGroup();
            this.IntakeTab.SuspendLayout();
            this.SuspendLayout();
            // 
            // IntakeTab
            // 
            this.IntakeTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.IntakeTab.Groups.Add(Test);
            this.IntakeTab.Label = "Email Intake";
            this.IntakeTab.Name = "IntakeTab";
            // 
            // Test
            // 
            Test.Label = "test";
            Test.Name = "Test";
            Test.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Test_DialogLauncherClick);
            // 
            // IntakeRibbon
            // 
            this.Name = "IntakeRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.IntakeTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon2_Load);
            this.IntakeTab.ResumeLayout(false);
            this.IntakeTab.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab IntakeTab;
    }

    partial class ThisRibbonCollection
    {
        internal IntakeRibbon Ribbon2
        {
            get { return this.GetRibbon<IntakeRibbon>(); }
        }
    }
}
