using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OEFCemail
{
    public partial class IntakeControl1 : UserControl
    {
        // property identifiers, used for attachment property checking
        // Reference for all identifies: https://interoperability.blob.core.windows.net/files/MS-OXPROPS/%5bMS-OXPROPS%5d.pdf
        // See 1.3.4.1 for how to use
        const string PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003";// See 2.594
        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";// See 2.1020
        const string SenderSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F";// See 2.1006

        private bool ProjectTextBoxActive = true; // for styling the project type radio buttons

        private Outlook.MailItem mailItem;
        
        public IntakeControl1()
        {
            InitializeComponent();
        }

        private void Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        #region Autofill
        /*
        private Outlook.MailItem GetMailItem()
        {
            Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            if (explorer != null)
            {
                Object obj = explorer.Selection[1];
                if (obj is Outlook.MailItem)
                    return (obj as Outlook.MailItem);
            }
            return null;
        }
        */
        public void SetMailItem(Outlook.MailItem mi)
        {
            mailItem = mi;
        }

        public void AutoFillFields()
        {
            this.textBoxSubject.Text = mailItem.Subject;
            this.textBoxSender.Text = mailItem.SenderName.ToString();

            // alternative to item.SenderEmailAddress, more reliable to getting in-office email addresses.
            string senderAddress = mailItem.PropertyAccessor.GetProperty(SenderSmtpAddress).ToString();
            this.textBoxSender.Text += " (" + senderAddress + ")";
            this.textBoxTime.Text = mailItem.ReceivedTime.ToString();

            FillRecipientsTextBox(mailItem);
            FillAttachmentsTextBox(mailItem);
        }

        private void FillRecipientsTextBox(Outlook.MailItem item)
        {
            // in the case the recipients have values in this, empty them
            this.textBoxReceiver.Text = "";

            Outlook.Recipients recip = item.Recipients; //includes CCs
            for (int i = 1; i <= recip.Count; i++)
            {
                Outlook.Recipient r = recip[i];
                this.textBoxReceiver.Text += r.Name;

                // https://docs.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-get-the-e-mail-address-of-a-recipient
                // alternative to r.Address, more reliable to getting in-office email addresses.
                string smtpAddress = r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS).ToString();
                this.textBoxReceiver.Text += " (" + smtpAddress + ")";
                if (i < recip.Count)
                    this.textBoxReceiver.Text += "; ";
            }
        }

        private void FillAttachmentsTextBox(Outlook.MailItem item)
        {
            // in the case the attachments/recipients have values in this, empty them
            this.textBoxAttach.Text = "";

            Outlook.Attachments attach = item.Attachments;
            for (int i = 1; i <= attach.Count; i++)
            {
                Outlook.Attachment att = attach[i];
                if (!IsEmbedded(att))
                    this.textBoxAttach.Text += att.FileName + ", ";
            }

            // If there are attachments, remove the final ", " at the end
            if (!this.textBoxAttach.Text.Equals(""))
                this.textBoxAttach.Text = this.textBoxAttach.Text.Remove(this.textBoxAttach.Text.Length - 2);
        }

        // check if attachment is embedded. Returns true if it is
        private bool IsEmbedded(Outlook.Attachment att) {
            Outlook.PropertyAccessor pa = att.PropertyAccessor;
            int flag = pa.GetProperty(PR_ATTACH_FLAGS);

            // https://stackoverflow.com/questions/3880346/dont-save-embed-image-that-contain-into-attachements-like-signature-image
            // https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcmsg/af8700bc-9d2a-47e4-b107-5ebf4467a418
            // flag of 4 -> the attachment is embedded in the message object's HTML body
            // Type = 6 -> Rich Text Format. This ensures not including embedded images, while still including attachments.
            if (flag != 4 && (int)att.Type != 6)
                return false;

            return true;
        }

        #endregion

        #region Save Email to File

        // only allow inputting project numbers for file lookup
        private void RadioButtonPrj_CheckedChanged(object sender, EventArgs e)
        {
            ProjectTextBoxActive = !ProjectTextBoxActive;
            if (ProjectTextBoxActive) {
                this.textBoxProject.BackColor = System.Drawing.Color.White;
                this.textBoxProject.ReadOnly = false;
            } else {
                this.textBoxProject.BackColor = System.Drawing.Color.Gray;
                this.textBoxProject.ReadOnly = true;
            }
                
        }

        private void ButtonSaveEmail_Click(object sender, EventArgs e)
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
                MessageBox.Show("Mail Item Not Selected.");
            }

        }
        #endregion

        #region Save Contents
        private void ButtonAppend_Click(object sender, EventArgs e)
        {
            if(mailItem != null)
            {
                string dir = GetProjectDirectory();
                string[] content =
                {
                    this.textBoxSubject.Text,
                    this.textBoxSender.Text,
                    this.textBoxReceiver.Text,
                    this.textBoxTime.Text,
                    this.textBoxAttach.Text
                };

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
                        EmailSaver emailSaver = new EmailSaver(openFileDialog.FileName, content, mailItem);
                        try
                        {
                            //TODO progress bar?
                            emailSaver.Save();
                        }
                        catch (Exception exc)
                        {
                            MessageBox.Show(exc + "\nError Saving to Word Doc. Suspending Process...");
                            emailSaver.SuspendProcess();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Mail Item Not Selected.");
            }
            
        }

        /*
        // check if required fields are empty. Display the empty fields and return true if any are empty
        private bool FieldsEmpty(string[] content)
        {
            bool empty = false;
            string emptyFields = "";
            for (int i = 0; i < content.Length - 1; i++)
            {
                if (content[i].Equals(""))
                {
                    //field separator
                    if (!emptyFields.Equals(""))
                        emptyFields += ", ";
                    else
                        empty = true;

                    switch (i)
                    {
                        case 0:
                            emptyFields += "Subject";
                            break;
                        case 1:
                            emptyFields += "Sender";
                            break;
                        case 2:
                            emptyFields += "Receiver";
                            break;
                        case 3:
                            emptyFields += "Time";
                            break;
                    }
                }
            }

            if(empty)
                MessageBox.Show("The following fields are empty:\n" +
                    emptyFields);

            return empty;
        }
        */
        #endregion

        #region Get Filepath
        // Return file path used for initial filepath for SaveFileDialog
        private String GetProjectDirectory()
        {
            string dir = "G:\\"; // the directory of the Projects drive using Windows formatting
            string prj = this.textBoxProject.Text;

            // Given a project #, go into the folder of the year that project is in (20 + the first 2 numbers of the given prj #
            //  and find a folder with the same project #
            if (this.radioButtonPrj.Checked && prj.Length > 1)
            {
                string path = dir + "20" + prj.Substring(0, 2) + " Projects\\";
                string s = SearchDirectories(path, prj);
                if (!s.Equals(""))
                    dir = s;
            }
            else if (this.radioButtonAR.Checked)
            {
                dir += "At Risk\\";
            }
            else if (this.radioButtonOH.Checked)
            {
                dir += "OverHead Projects (OHPs)\\";
            }

            return dir;
        }

        private string SearchDirectories(string path, string prj)
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
                MessageBox.Show("Invalid directory. Check if the Project # is correct.");
                Console.WriteLine(e);
            }
            return "";
        }
        #endregion
    }
}
