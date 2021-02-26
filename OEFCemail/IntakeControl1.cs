using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Outlook;

namespace OEFCemail
{
    public partial class IntakeControl1 : UserControl
    {
        // property identifier, used for attachment property checking
        // See 1.3.4.1 and 2.587 here: https://interoperability.blob.core.windows.net/files/MS-OXPROPS/%5bMS-OXPROPS%5d.pdf
        const string PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";
        const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        private bool ProjectTextBoxActive = true;
        
        public IntakeControl1()
        {
            InitializeComponent();
        }

        private void Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private Outlook.MailItem GetMailItem()
        {
            // https://codesteps.com/2018/08/06/outlook-2010-add-in-get-mailitem-using-c/
            // Get currently selected Outlook item on button click. If mailitem, parse the necessary fields.
            Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            if (explorer != null)
            {
                Object obj = explorer.Selection[1];
                if (obj is Outlook.MailItem)
                    return (obj as Outlook.MailItem);
            }
            return null;
        }

        private void ButtonAutoFill_Click(object sender, EventArgs e)
        {
            Outlook.MailItem item = GetMailItem();
            if (item != null)
            {
                //TODO: parse content for better formatting?
                this.textBoxSender.Text = item.SenderName.ToString();
                this.textBoxTime.Text = item.ReceivedTime.ToString();
                this.textBoxContent.Text = item.Body.ToString();

                // in the case the attachments/recipients have values in this, empty them
                this.textBoxReceiver.Text = "";
                this.textBoxAttach.Text = "";

                FillRecipientsTextBox(item);
                FillAttachmentsTextBox(item);
            }

        }

        private void FillRecipientsTextBox(Outlook.MailItem item)
        {
            Outlook.Recipients recip = item.Recipients; //includes CCs
            for (int i = 1; i <= recip.Count; i++)
            {
                Outlook.Recipient r = recip[i];
                //TODO: fix sizing for text boxes
                this.textBoxReceiver.Text += r.Name;

                // https://docs.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-get-the-e-mail-address-of-a-recipient
                string smtpAddress = r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS).ToString();
                this.textBoxReceiver.Text += " (" + smtpAddress + ")";
                if (i < recip.Count)
                    this.textBoxReceiver.Text += "; ";
            }
        }

        private void FillAttachmentsTextBox(Outlook.MailItem item)
        {
            //TODO: fix attachment parsing
            Outlook.Attachments attach = item.Attachments;
            for (int i = 1; i <= attach.Count; i++)
            {
                Outlook.Attachment att = attach[i];
                if (!IsEmbedded(att))
                {
                    this.textBoxAttach.Text += att.FileName;
                    if (i < attach.Count)
                        this.textBoxAttach.Text += ", ";
                }

            }
        }

        // https://stackoverflow.com/questions/59075501/find-out-if-an-attachment-is-embedded-or-attached
        // check if attachment is embedded. Returns true if it is
        private bool IsEmbedded(Outlook.Attachment att) {
            string s = "";
            try {
                s = (string)att.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID);
            } catch(Exception e) {
                MessageBox.Show("Error getting attachment property PR_ATTACH_CONTENT_ID.");
                Console.Write(e);
            }
            
            if (s == "")
                return false;

            return true;
        }

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
            Outlook.MailItem item = GetMailItem();
            //TODO parse OEFC specific emails
            String dir = GetProjectDirectory();

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Outlook Message File|*.msg",
                Title = "Save an Email",
                RestoreDirectory = true,
                InitialDirectory = dir
            };
            saveFileDialog.ShowDialog();

            // If the file name is not an empty string open it for saving.
            //TODO parse mailitem body to trim email down as needed
            if (saveFileDialog.FileName != "")
                item.SaveAs(saveFileDialog.FileName, Outlook.OlSaveAsType.olMSG);

        }

        private void ButtonAppend_Click(object sender, EventArgs e)
        {
            String dir = GetProjectDirectory();

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
                // https://www.codeproject.com/Questions/1104176/Append-text-to-existing-word-document-from-another
                //TODO formatting (include sender/receiver/content/attachments/date)
                //TODO progress bar?
                //TODO only save to Project Notes?
                //TODO parsing content for less editing?
                //TODO parsing sender/receiver?
                Word._Application oWord = new Word.Application();
                oWord.Documents.Open(openFileDialog.FileName);
                oWord.Selection.TypeText(this.textBoxContent.Text);
                oWord.ActiveDocument.Save();
                oWord.Quit();
            }
        }

        // Return file path used for initial filepath for SaveFileDialog
        private String GetProjectDirectory()
        {
            String dir = "G:\\";
            String prj = this.textBoxProject.Text;

            if (this.radioButtonPrj.Checked && !prj.Equals("") && prj.Length > 1)
            {
                String path = dir + "20" + prj.Substring(0, 2) + " Projects\\";
                String s = SearchDirectories(path, prj);
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

        private String SearchDirectories(String path, String prj)
        {
            try
            {

                if (Directory.EnumerateDirectories(path, prj + "*").Any())
                {
                    String[] s = Directory.GetDirectories(path, prj + "*");
                    return s[0];
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Invalid directory. Check if the Project # is correct.");
                Console.Write(e);
            }
            return "";
        }
 
    }
}
