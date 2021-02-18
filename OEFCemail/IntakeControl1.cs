using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace OEFCemail
{
    public partial class IntakeControl1 : UserControl
    {

        private bool ProjectTextBoxActive = true;

        public IntakeControl1()
        {
            InitializeComponent();
        }

        private void Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void ButtonAutoFill_Click(object sender, EventArgs e)
        {
            Outlook.MailItem item = GetMailItem();
            if (item != null)
            {
                //TODO: parsing sender/receiver?
                //TODO: parse content for better formatting?
                this.textBoxSender.Text = item.SenderName.ToString();
                this.textBoxReceiver.Text = item.ReceivedByName;
                this.textBoxTime.Text = item.ReceivedTime.ToString();
                this.textBoxContent.Text = item.Body.ToString();
                //TODO: parse attachment names (first check if attachments, then parse through list)
                //this.textBoxAttach.Text =
            }

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
