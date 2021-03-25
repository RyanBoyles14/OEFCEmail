using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
        private string emailBody;
        private readonly Word._Application oWord;
        private readonly Word._Document oDoc;
        public EmailSaver(string filename, string[] content)
        {
            this.filename = filename;
            subject = content[0];
            sender = content[1];
            receiver = content[2];
            time = content[3];
            emailBody = content[4];

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
            bool hasMoreMessages = true; 
            string sub = TrimSubject(subject);
            DateTime dt = ParseTime(time); // time for the top message
            while (hasMoreMessages) {
                // i=0: beginning index of the next message
                // i=1: length of the email properties to parse through
                int[] propertyIndices = GetSegmentInfo();
                if (propertyIndices[0] == -1)
                    hasMoreMessages = false;

                string msg = ParseContents(propertyIndices[0], propertyIndices[1]);
                MessageBox.Show(msg);

                int row = FindRow(sub, dt);
                if (row == -1)
                {
                    MessageBox.Show("Current message may have already been saved. Suspending the process.");
                    break;
                }

                // write to Word Doc

                // prepare for next cycle by getting the next message's properties.
                if(hasMoreMessages)
                {
                    ParseNextMessageProperties(propertyIndices[1]);
                }
            }
            

            oWord.Quit(); 
        }

        #region Trim/Parse
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
        private DateTime ParseTime(string dt)
        {
            string pattern;
            DateTime parsedDate;
            if (dt.StartsWith("Sent: ")) {
                // Format from contents: "Sent: Day, Month xx, 20xx x:xx PM"
                // https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings
                pattern = "Sent: dddd, M d, yyyy h:mm tt";
            } else {
                // Format from textBoxTime.Text: "x/xx/20xx x:xx:xx xM"
                pattern = "M/d/yyyy h:mm:ss tt";
            }

            parsedDate = DateTime.ParseExact(dt, pattern, null, System.Globalization.DateTimeStyles.None);
            return parsedDate;
        }

        /*
         * Using Regex, find other messages down the chain
         * Using this format to find other messages in a chain
         * From: X
         * Sent: X
         * To: X
         * Cc: X - this is optional
         * Subject: X
         */
        private int[] GetSegmentInfo()
        {
            string rExp = "(From: .{0,}\n)" +
                "(Sent: .{0,}\n)" +
                "(To: .{0,}\n)" +
                "(Cc: .{0,}\n){0,1}" +
                "(Subject: .{0,}\n)";
            Match match = Regex.Match(emailBody, rExp);

            int propertyIndex = match.Success ? match.Index : -1;

            int[] propertyIndices = { propertyIndex, match.Length };
            return propertyIndices;
        }

        //TODO parse/formatting (include sender/receiver/content/attachments/timestamp/subject)
        //TODO trim whitespace?
        private string ParseContents(int segmentIndex, int msgInfoLength)
        {
            //TODO separate messages from threads
            //TODO find cut-off for threads
            string msg;
            if (segmentIndex == -1)
            {
                // no other messages in the chain, set to what's left of the message
                msg = emailBody;
            }
            else
            {
                // separate the latest message from the chain, and remove from messageBody.
                msg = emailBody.Substring(0, segmentIndex);
                emailBody = emailBody.Remove(0, segmentIndex);
            }

            return msg;
        }

        private string ParseNextMessageProperties(int length)
        {
            emailBody = emailBody.Remove(0, length);
            return "";
        }
        #endregion

        #region Find Rows
        //TODO steamline this? Compare Ranges?
        private int FindRow(string sub, DateTime dt)
        {
            int row = 0;
            Word.Table oTbl = oDoc.Tables[1];
            Word.Range rng = oTbl.Range; // Assuming there is only one table in the project notes
            object findSub = "[Subject: " + sub + "]"; // Using the new note format "[Subject: X]"

            // If subject is found in the table, then find the row.
            // Else, we can assume we can put in the last row
            if (rng.Find.Execute(ref findSub))
            {
                oTbl.Columns[1].Select();
                // Search all rows, from the bottom up (most recent), for any with the current mail subject
                
                for (int i = oTbl.Rows.Count; i > 0; i--)
                {
                    Word.Range rowRng = oTbl.Rows[i].Range;
                    rowRng.Find.ClearFormatting();

                    Word.Range subjectRng = rowRng.Sentences[1];
                    Word.Range timeRng = rowRng.Sentences[2];

                    if (subjectRng.Text.TrimEnd('\n').CompareTo((string)findSub) == 0)
                    {
                        int result = CompareDates(timeRng.Text.TrimEnd('\n'), dt);
                        if(result < 0) { //The current row has an earlier timestamp
                            row = i;
                            break;
                        } else if(result == 0) { //Usually means the current notes are already intaken.
                            row = -1;
                        }
                    }
                    
                }
            }

            return row;
        }

        private int CompareDates(string t, DateTime dt)
        {
            string pattern = "[Time: MM-dd-yy h:mmtt]"; // Using the new note format
            DateTime parsedDate = DateTime.ParseExact(t, pattern, null, System.Globalization.DateTimeStyles.None);

            return DateTime.Compare(parsedDate, dt);
        }

        #endregion 

        // Insert into Doc.
        private void InsertInDoc()
        {
            //date time .ToString("M/d/yy hh:mmtt")
        }
    }
}
