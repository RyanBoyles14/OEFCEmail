﻿using System;
using static OEFCemail.ErrorLog;
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

            // set new ErrorLog for IntakeRibbon to use.
            Globals.Ribbons.IntakeRibbon.ErrorLog = new ErrorLog();
            // Create event for when error log is triggered (when user wants to send error report)
            Globals.Ribbons.IntakeRibbon.ErrorLog.SendErrorReport += ErLog_SendErrorReport;
        }

        // When the ErrorLog is used and the user wishes to send an error report, create an email.
        private void ErLog_SendErrorReport(object sender, SendErrorReportEventArgs args)
        {
            Outlook.MailItem email = (Outlook.MailItem)this.Application.CreateItem(Outlook.OlItemType.olMailItem);
            email.Subject = "OEFCemail Exception: " + args.Time;
            email.Body = "";
            email.To = "OEFCEmail@gmail.com";
            email.Importance = Outlook.OlImportance.olImportanceLow;
            
            try
            {
                email.Attachments.Add(args.File, Outlook.OlAttachmentType.olByValue);
                ((Outlook._MailItem)email).Send();
            }
            catch
            {
                email.Delete();
            }
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

        // When the user selects a different mail item,
        // update the mail item currently selected in the IntakeRibbon,
        // and change the labels showing the currently selected item
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
