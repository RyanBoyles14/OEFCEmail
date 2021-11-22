using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using JR.Utils.GUI.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OEFCemail
{
    public partial class IntakeRibbon
    {

        // property identifiers, used for attachment property checking
        // Reference for all identifies: https://interoperability.blob.core.windows.net/files/MS-OXPROPS/%5bMS-OXPROPS%5d.pdf
        // See 1.3.4.1 for how to use
        private const string PidTagAttachFlags = "http://schemas.microsoft.com/mapi/proptag/0x37140003";// See 2.594
        private const string PidTagSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";// See 2.1020
        private const string PidTagSenderSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F";// See 2.1006

        private Outlook.MailItem mailItem;

        private void IntakeRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        public void SetMailItem(Outlook.MailItem mi)
        {
            mailItem = mi;
        }

        private string GetSender()
        {
            return mailItem.SenderName.ToString() + " (" + mailItem.PropertyAccessor.GetProperty(PidTagSenderSmtpAddress).ToString() + ")";
        }

        private string GetRecipients()
        {
            string s = "";

            Outlook.Recipients recip = mailItem.Recipients; //includes CCs
            for (int i = 1; i <= recip.Count; i++)
            {
                Outlook.Recipient r = recip[i];
                s += r.Name;

                // https://docs.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-get-the-e-mail-address-of-a-recipient
                // alternative to r.Address, more reliable to getting in-office email addresses.
                string smtpAddress = r.PropertyAccessor.GetProperty(PidTagSmtpAddress).ToString();
                s += " (" + smtpAddress + ")";
                if (i < recip.Count)
                    s += "; ";
            }

            return s;
        }

        private string GetAttachments()
        {
            string s = "";

            Outlook.Attachments attach = mailItem.Attachments;
            for (int i = 1; i <= attach.Count; i++)
            {
                Outlook.Attachment att = attach[i];
                if (!IsEmbedded(att))
                {
                    s += att.FileName;
                    if (i < attach.Count)
                        s += ", ";
                }  
            }

            return s;
        }

        // check if attachment is embedded. Returns true if it is
        private static bool IsEmbedded(Outlook.Attachment att)
        {
            Outlook.PropertyAccessor pa = att.PropertyAccessor;
            int flag = pa.GetProperty(PidTagAttachFlags);

            // https://stackoverflow.com/questions/3880346/dont-save-embed-image-that-contain-into-attachements-like-signature-image
            // https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcmsg/af8700bc-9d2a-47e4-b107-5ebf4467a418
            // flag of 4 -> the attachment is embedded in the message object's HTML body
            // Type = 6 -> Rich Text Format. This ensures not including embedded images, while still including attachments.
            if (flag != 4 && (int)att.Type != 6)
                return false;

            return true;
        }

        private String GetProjectDirectory()
        {
            string dir = "G:\\"; // the directory of the Projects drive using Windows formatting
            string folder = folderLocationDropDown.SelectedItem.ToString();
            string prj = projectEditBox.Text;
            string path = dir;
            string s;

            switch (folder)
            {
                case "Projects":
                    if (prj.Length > 1)
                        // Given a project #, go into the folder of the year
                        // that project is in (20 + the first 2 numbers of the given prj #)
                        // and find a folder with the same project #
                        path += "20" + prj.Substring(0, 2) + " Projects\\";

                    break;
                case "At Risk":
                    path += "At Risk\\";
                    break;
                case "Overhead":
                    path += "OverHead Projects (OHPs)\\";
                    break;
            }

            if (prj.Length > 1)
            {
                s = SearchDirectories(path, prj);
                if (s.Equals(""))
                    dir = path;
                else
                    dir = s;
            }
            else
                dir = path;

            return dir;
        }

        private static string SearchDirectories(string path, string prj)
        {
            try
            {
                if (Directory.EnumerateDirectories(path, prj + "*").Any())
                {
                    string[] s = Directory.GetDirectories(path, prj + "*");
                    return s[0];
                }
            }
            catch (Exception e)
            {
                FlexibleMessageBox.Show("Couldn't find selected directory.\n");
                Console.Write(e);
            }
            return "";
        }

        private void SaveEmailToFileButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (mailItem != null)
            {
                string dir = GetProjectDirectory();

                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Outlook Message File|*.msg",
                    Title = "Save an Email",
                    RestoreDirectory = true,
                    InitialDirectory = dir,
                    FileName = mailItem.Subject
                };
                saveFileDialog.ShowDialog();

                // If the file name is not an empty string open it for saving.
                if (saveFileDialog.FileName != "")
                    mailItem.SaveAs(saveFileDialog.FileName, Outlook.OlSaveAsType.olMSG);
            }
            else
            {
                FlexibleMessageBox.Show("Mail Item Not Selected.");
            }
        }


        private async void SaveEmailToNotesButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (mailItem != null)
            {
                string sub;
                if (mailItem.Subject == null)
                    sub = "(no subject)";
                else
                    sub = mailItem.Subject;

                string dir = GetProjectDirectory();

                if (mailItem != null)
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog
                    {
                        Filter = "Word Documents|*.docx",
                        Title = "Open a Word Doc File",
                        RestoreDirectory = true,
                        InitialDirectory = dir
                    };
                    openFileDialog.ShowDialog();

                    if (openFileDialog.FileName != "")
                    {
                        EmailSaver emailSaver = new EmailSaver(openFileDialog.FileName, sub,
                            GetSender(), GetRecipients(), mailItem.ReceivedTime.ToString(), GetAttachments());

                        if (emailSaver.Initialized)
                        {
                            //Make sure to test with read-only and not "selecting" a mail item
                            try
                            {
                                await Task.Run(() => emailSaver.SaveAsync(mailItem));
                            }
                            catch (Exception exc)
                            {
                                FlexibleMessageBox.Show(exc.Message);

                                ErrorLog log = new ErrorLog();
                                log.WriteErrorLog(exc.ToString());

                                emailSaver.TerminateProcess();
                            }
                        }
                    }
                }
            }
            else if (mailItem == null)
            {
                FlexibleMessageBox.Show("Mail Item Not Selected.");
            }
            else
            {
                FlexibleMessageBox.Show("Mail Item Does Not Have a Subject.");
            }
        }
    }
}
