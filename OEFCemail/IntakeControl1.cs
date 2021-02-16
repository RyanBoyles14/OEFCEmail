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

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void buttonAutoFill_Click(object sender, EventArgs e)
        {

            // https://codesteps.com/2018/08/06/outlook-2010-add-in-get-mailitem-using-c/
            // Get currently selected Outlook item on button click. If mailitem, parse the necessary fields.
            Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            if (explorer != null)
            {
                Object obj = explorer.Selection[1];
                if (obj is Outlook.MailItem)
                {
                    Outlook.MailItem item = (obj as Outlook.MailItem);

                    //TODO parse OEFC specific emails
                    this.textBoxSender.Text = item.SenderName.ToString();
                    this.textBoxReceiver.Text = item.ReceivedByName; 
                    this.textBoxTime.Text = item.ReceivedTime.ToString();
                    this.textBoxContent.Text = item.Body.ToString();
                    //TODO: parse attachment names (first check if attachments, then parse through list)
                    //this.textBoxAttach.Text =
                }
            }

        }
    }
}
