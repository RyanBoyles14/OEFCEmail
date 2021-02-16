
namespace OEFCemail
{
    partial class IntakeRibbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public IntakeRibbon1()
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
            this.IntakeTab1 = this.Factory.CreateRibbonTab();
            Test = this.Factory.CreateRibbonGroup();
            this.IntakeTab1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Test
            // 
            Test.Label = "Show User Control";
            Test.Name = "Test";
            Test.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Test_DialogLauncherClick);
            // 
            // IntakeTab1
            // 
            this.IntakeTab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.IntakeTab1.Groups.Add(Test);
            this.IntakeTab1.Label = "Email Intake";
            this.IntakeTab1.Name = "IntakeTab1";
            // 
            // IntakeRibbon1
            // 
            this.Name = "IntakeRibbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.IntakeTab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.IntakeRibbon_Load);
            this.IntakeTab1.ResumeLayout(false);
            this.IntakeTab1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab IntakeTab1;
    }

    partial class ThisRibbonCollection
    {
        internal IntakeRibbon1 Ribbon2
        {
            get { return this.GetRibbon<IntakeRibbon1>(); }
        }
    }
}
