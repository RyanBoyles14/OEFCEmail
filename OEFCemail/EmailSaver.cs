using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Outlook;

namespace OEFCemail
{
    class EmailSaver
    {
        private readonly string filename;
        private readonly string subject;
        private readonly string sender;
        private readonly string receiver;
        private readonly string time;
        private readonly string message;
        private readonly Word._Application oWord;
        private readonly Word._Document oDoc;
        public EmailSaver(string filename, string[] content)
        {
            this.filename = filename;
            subject = content[0];
            sender = content[1];
            receiver = content[2];
            time = content[3];
            message = content[4];

            oWord = new Word.Application();

            try
            {
                oDoc = oWord.Documents.Open(filename);
            }
            catch (Exception e)
            {
                if (e is IOException)
                    MessageBox.Show("Error Opening Word Doc. Check that it is not already open");
                Console.WriteLine(e);
            }
        }
        public void Save()
        {
            //TODO progress bar?
            //TODO append formatted content at correct spot
            //TODO ensure embedded images and links get included in project notes
            /*
            if (!oDoc.ReadOnly) // user can still open the file, but the program cannot save to it
            {
                Word.Tables tables = oDoc.Tables;
                foreach(Word.Table table in tables) { //in the case there are multiple tables
                    table.Cell(1, 1).Range.Text = 
                        content[0] + "\n" + //subject
                        content[3] + "\n" + //time
                        content[4] + "\n" + //contents
                        "(Attachment:" + content[5] + ")"; //attachments

                    table.Cell(1, 2).Range.Text = content[1] + " to " + content[2]; //sender to receiver
                    //oWord.Selection.TypeText(this.textBoxContent.Text);
                    oWord.ActiveDocument.Save();
                }
            }

            */
            string sub = TrimSubject(subject);
            string t = TrimTime(time);
            FindRow(subject, time);
            oWord.Quit();
            
        }

        // Get the base subject header w/o Forward or Reply prefixes
        public string TrimSubject(string sub)
        {
            // Typical Prefixes are "FW", "Fwd", or "RE". Search for any variation of those.
            while(sub.ToUpper().StartsWith("FW:") || sub.ToUpper().StartsWith("FWD:") || sub.ToUpper().StartsWith("RE:"))
            {
                
                if(sub.ToUpper().StartsWith("FWD:"))
                    sub = sub.Remove(0, 4);
                else
                    sub = sub.Remove(0, 3);

                // Usually, the prefixes have a space immediately proceeding. There may be abnormal situations where this isn't the case
                // Trim any spaces at the start of the header, in case of any atypical cases
                sub = sub.TrimStart(' ');
            }

            return sub;
        }
       
        // parse the time based on the typical formats from Outlook emails
        private string TrimTime(string t)
        {
            string pattern;
            DateTime parsedDate;
            if (t.StartsWith("Sent: ")) {
                // Format from contents: "Sent: Day, Month xx, 20xx x:xx PM"
                // https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings
                pattern = "Sent: dddd, M d, yyyy h:mm tt";
            } else {
                // Format from textBoxTime.Text: "x/xx/20xx x:xx:xx xM"
                pattern = "M/d/yyyy h:mm:ss tt";
            }

            parsedDate = DateTime.ParseExact(t, pattern, null, System.Globalization.DateTimeStyles.None);
            return parsedDate.ToString("M/d/yy hh:mm tt");
        }

        //TODO parsing contents of project notes to figure out where to insert contents/if contents already exist in the notes
        //TODO open and parse project notes
        private void FindRow(string sub, string t)
        {

        }

        //TODO parsing contents (by timestamp + subject for now) to find how much of the email thread needs saved
        //TODO separate messages from threads
        //TODO find cut-off for threads
        private void ParseSegment()
        {

        }

        //TODO formatting (include sender/receiver/content/attachments/timestamp/subject)
        private void ParseContents()
        {
        }

        //TODO trim subject as needed
        //TODO trim down excessive whitespace
        private string TrimContents()
        {
            return "";
        }

        // Insert into Doc.
        private void InsertInDoc()
        {

        }
    }
}
