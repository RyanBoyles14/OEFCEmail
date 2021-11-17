
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            this.Tab1 = this.Factory.CreateRibbonTab();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.folderLocationDropDown = this.Factory.CreateRibbonDropDown();
            this.projectEditBox = this.Factory.CreateRibbonEditBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.saveEmailToFileButton = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.saveEmailToNotesButton = this.Factory.CreateRibbonButton();
            this.Tab1.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // Tab1
            // 
            this.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Tab1.Groups.Add(this.group4);
            this.Tab1.Label = "Email Intake";
            this.Tab1.Name = "Tab1";
            // 
            // group4
            // 
            this.group4.Items.Add(this.folderLocationDropDown);
            this.group4.Items.Add(this.projectEditBox);
            this.group4.Items.Add(this.separator1);
            this.group4.Items.Add(this.saveEmailToFileButton);
            this.group4.Items.Add(this.separator2);
            this.group4.Items.Add(this.saveEmailToNotesButton);
            this.group4.Label = "Save Email";
            this.group4.Name = "group4";
            // 
            // folderLocationDropDown
            // 
            ribbonDropDownItemImpl1.Label = "Projects";
            ribbonDropDownItemImpl2.Label = "At Risk";
            ribbonDropDownItemImpl3.Label = "Overhead";
            this.folderLocationDropDown.Items.Add(ribbonDropDownItemImpl1);
            this.folderLocationDropDown.Items.Add(ribbonDropDownItemImpl2);
            this.folderLocationDropDown.Items.Add(ribbonDropDownItemImpl3);
            this.folderLocationDropDown.Label = "Folder Location";
            this.folderLocationDropDown.Name = "folderLocationDropDown";
            // 
            // projectEditBox
            // 
            this.projectEditBox.Label = "Project #";
            this.projectEditBox.Name = "projectEditBox";
            this.projectEditBox.Text = null;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // saveEmailToFileButton
            // 
            this.saveEmailToFileButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.saveEmailToFileButton.Image = global::OEFCemail.Properties.Resources.savefile;
            this.saveEmailToFileButton.Label = "Save Email To File";
            this.saveEmailToFileButton.Name = "saveEmailToFileButton";
            this.saveEmailToFileButton.ShowImage = true;
            this.saveEmailToFileButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveEmailToFileButton_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // saveEmailToNotesButton
            // 
            this.saveEmailToNotesButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.saveEmailToNotesButton.Image = global::OEFCemail.Properties.Resources.savenotes;
            this.saveEmailToNotesButton.Label = "Save Email to Notes";
            this.saveEmailToNotesButton.Name = "saveEmailToNotesButton";
            this.saveEmailToNotesButton.ShowImage = true;
            this.saveEmailToNotesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveEmailToNotesButton_Click);
            // 
            // IntakeRibbon
            // 
            this.Name = "IntakeRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.Tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.IntakeRibbon_Load);
            this.Tab1.ResumeLayout(false);
            this.Tab1.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown folderLocationDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox projectEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveEmailToFileButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveEmailToNotesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
    }

    partial class ThisRibbonCollection
    {
        internal IntakeRibbon IntakeRibbon
        {
            get { return this.GetRibbon<IntakeRibbon>(); }
        }
    }
}
