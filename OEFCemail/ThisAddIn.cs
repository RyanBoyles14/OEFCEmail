using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OEFCemail
{

    public partial class ThisAddIn
    {
        Outlook.Application app;
        Outlook.Explorer activeExplorer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Initialize the ActiveExplorer, which is used to for selecting mail items
            app = Application;
            activeExplorer = app.ActiveExplorer();

            // create an event handler for if the user selects a new item
            activeExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(Item_SelectionChange);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            // must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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

        void Item_SelectionChange()
        {
            // https://codesteps.com/2018/08/06/outlook-2010-add-in-get-mailitem-using-c/
            // Get currently selected Outlook item on button click. If mailitem, parse the necessary fields.
            if (activeExplorer != null && activeExplorer.Selection.Count > 0)
            {
                Object obj = activeExplorer.Selection[1];

                IntakeRibbon ir = Globals.Ribbons.IntakeRibbon;

                if (obj is Outlook.MailItem item)
                {
                    ir.SetMailItem(item);

                    if (item.Subject == null)
                        ir.emailLabel.Label = "Subject: (no subject)";
                    else 
                        ir.emailLabel.Label = "Subject: " + item.Subject;

                    ir.senderLabel.Label = "Sender: " + item.Sender.Name;
                    ir.dateLabel.Label = "Date: " + item.ReceivedTime.ToString();
                }
            }
        }
    }
}
