using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OEFCemail
{
    public partial class IntakeControl1 : UserControl
    {
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
            //TODO parse OEFC specific emails
            if (item != null)
            {
                this.textBoxSender.Text = item.SenderName.ToString();
                this.textBoxReceiver.Text = item.ReceivedByName;
                this.textBoxTime.Text = item.ReceivedTime.ToString();
                this.textBoxContent.Text = item.Body.ToString();
                //TODO: parse attachment names (first check if attachments, then parse through list)
                //this.textBoxAttach.Text =
            }

        }

        private void ButtonSaveEmail_Click(object sender, EventArgs e)
        {

            Outlook.MailItem item = GetMailItem();
            //TODO parse OEFC specific emails
            if (item != null)
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog
                {
                    Filter = "Outlook Message File|*.msg",
                    Title = "Save an Email"
                };
                saveFileDialog1.ShowDialog();

                // If the file name is not an empty string open it for saving.
                if (saveFileDialog1.FileName != "")
                    item.SaveAs(saveFileDialog1.FileName, Outlook.OlSaveAsType.olMSG);
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
    }
}
