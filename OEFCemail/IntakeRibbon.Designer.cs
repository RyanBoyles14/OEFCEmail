
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(IntakeRibbon));
            this.Tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.folderLocationDropDown = this.Factory.CreateRibbonDropDown();
            this.projectEditBox = this.Factory.CreateRibbonEditBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.saveEmailToFileButton = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.saveEmailToNotesButton = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.emailLabel = this.Factory.CreateRibbonLabel();
            this.senderLabel = this.Factory.CreateRibbonLabel();
            this.dateLabel = this.Factory.CreateRibbonLabel();
            this.Tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // Tab1
            // 
            this.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Tab1.Groups.Add(this.group1);
            this.Tab1.Groups.Add(this.group2);
            this.Tab1.Label = "Email Intake";
            this.Tab1.Name = "Tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.folderLocationDropDown);
            this.group1.Items.Add(this.projectEditBox);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.saveEmailToFileButton);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.saveEmailToNotesButton);
            this.group1.Label = "Save Email";
            this.group1.Name = "group1";
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
            this.saveEmailToFileButton.Image = ((System.Drawing.Image)(resources.GetObject("saveEmailToFileButton.Image")));
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
            this.saveEmailToNotesButton.Image = ((System.Drawing.Image)(resources.GetObject("saveEmailToNotesButton.Image")));
            this.saveEmailToNotesButton.Label = "Save Email to Notes";
            this.saveEmailToNotesButton.Name = "saveEmailToNotesButton";
            this.saveEmailToNotesButton.ShowImage = true;
            this.saveEmailToNotesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveEmailToNotesButton_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.emailLabel);
            this.group2.Items.Add(this.senderLabel);
            this.group2.Items.Add(this.dateLabel);
            this.group2.Label = "Selected Email";
            this.group2.Name = "group2";
            // 
            // emailLabel
            // 
            this.emailLabel.Label = "Subject: ";
            this.emailLabel.Name = "emailLabel";
            // 
            // senderLabel
            // 
            this.senderLabel.Label = "Sender: ";
            this.senderLabel.Name = "senderLabel";
            // 
            // dateLabel
            // 
            this.dateLabel.Label = "Date: ";
            this.dateLabel.Name = "dateLabel";
            // 
            // IntakeRibbon
            // 
            this.Name = "IntakeRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.Tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.IntakeRibbon_Load);
            this.Tab1.ResumeLayout(false);
            this.Tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown folderLocationDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox projectEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveEmailToFileButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveEmailToNotesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel emailLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel senderLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel dateLabel;
    }

    partial class ThisRibbonCollection
    {
        internal IntakeRibbon IntakeRibbon
        {
            get { return this.GetRibbon<IntakeRibbon>(); }
        }
    }
}
