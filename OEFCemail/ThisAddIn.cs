using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OEFCemail
{

    public partial class ThisAddIn
    {
        //global variables to avoid the garabage collector
        private IntakeControl1 myIntakeControl1;
        private Microsoft.Office.Tools.CustomTaskPane intakeTaskPane;
        Outlook.Application app;
        Outlook.Explorer activeExplorer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Initialize the Side TaskPane
            myIntakeControl1 = new IntakeControl1();
            intakeTaskPane = this.CustomTaskPanes.Add(myIntakeControl1, "OEFC Email Saver");
            // Create EventHandler to handle the pane's visibility
            intakeTaskPane.VisibleChanged += new EventHandler(IntakeTaskPane_VisibleChanged);

            // Initialize the ActiveExplorer, which is used to for selecting mail items
            app = this.Application;
            activeExplorer = app.ActiveExplorer(); 
            // create an event handler for if the user selects a new item
            activeExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(Item_SelectionChange);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        private void IntakeTaskPane_VisibleChanged(object sender, System.EventArgs e)
        {
            Globals.Ribbons.IntakeRibbon.toggleButtonIntakeDisplay.Checked =
                intakeTaskPane.Visible;
        }

        void Item_SelectionChange()
        {
            // https://codesteps.com/2018/08/06/outlook-2010-add-in-get-mailitem-using-c/
            // Get currently selected Outlook item on button click. If mailitem, parse the necessary fields.
            if (activeExplorer != null && activeExplorer.Selection.Count > 0)
            {
                Object obj = activeExplorer.Selection[1];
                if (obj is Outlook.MailItem item)
                {
                    myIntakeControl1.SetMailItem(item);
                    myIntakeControl1.AutoFillFields();
                }
            }
        }

        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get
            {
                return intakeTaskPane;
            }
        }
    }
}
