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
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OEFCemail
{
    class EmailSaver
    {
        private readonly string filename;
        private readonly string subject;
        private readonly string sender;
        private readonly string receiver;
        private readonly string time;
        private readonly string attachment;
        private readonly Word._Application oWord;
        private readonly Word._Document oDoc;
        private readonly Word.Document mailInspector;

        public EmailSaver(string filename, string[] content, Outlook.MailItem item)
        {
            this.filename = filename;
            subject = content[0];
            sender = content[1];
            receiver = content[2];
            time = content[3];
            attachment = content[4];

            oWord = new Word.Application();

            try
            {
                oDoc = oWord.Documents.Open(filename);
                mailInspector = item.GetInspector.WordEditor as Word.Document;
                mailInspector.Unprotect();
            }
            catch (Exception e)
            {
                if (e is IOException)
                    MessageBox.Show("Error Opening Word Doc. Check that it is not already open");
                Console.WriteLine(e);
            }            
        }

        //TODO progress bar?
        //TODO trim down whitespace/signoffs?
        //TODO test exiting loop
        public void Save()
        {
            bool success = true;
            if (oDoc != null)
            {
                if (!oDoc.ReadOnly) // user can still open the file, but the program cannot save to it
                {
                    bool hasMoreMessages = true;
                    string sub = TrimSubject(subject);
                    string send = sender;
                    string rec = receiver;
                    DateTime dt = ParseTime(time, false); // time for the top message
                    while (hasMoreMessages)
                    {
                        // index 0: beginning index of the next message
                        // index 1: length of the email properties to parse through
                        int[] propRanges = GetSegmentInfo();
                        if (propRanges[0] == -1)
                            hasMoreMessages = false;

                        ExtractMessage(propRanges[0]); ;

                        int row = FindRow(sub, dt);
                        if (row == -1)
                        {
                            MessageBox.Show("Current message may have already been saved. Suspending the process.");
                            success = false;
                            break;
                        }

                        // write to Word Doc
                        InsertInDoc(sub, send, rec, dt.ToString("MM-dd-yy h:mmtt"), row);

                        // prepare for next cycle by getting the next message's properties.
                        if (hasMoreMessages)
                        {
                            string[] prop = ParseNextMessageProperties(propRanges[1]);
                            send = prop[0];
                            dt = ParseTime(prop[1], true);
                            rec = prop[2];
                        }
                    }
                    if (success)
                        oDoc.ActiveWindow.Visible = true;
                    else
                        oWord.Quit();
                }
            }
            else
            {
                oWord.Quit();
                MessageBox.Show("Word Doc not opened. Suspending Process...");
            }
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
        private DateTime ParseTime(string t, bool moreMsg)
        {
            string pattern;
            DateTime parsedDate;
            if (moreMsg) {
                // Using the Format from emails in the chain: "Day, Month d, yyyy h:mm xM"
                // https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings
                pattern = "dddd, MMMM d, yyyy h:mm tt";
            } else {
                // Format from textBoxTime.Text: "x/xx/20xx x:xx:xx xM"
                pattern = "M/d/yyyy h:mm:ss tt";
            }

            parsedDate = DateTime.ParseExact(t, pattern, null, System.Globalization.DateTimeStyles.AssumeLocal);
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
            int startRange = -1;
            int length = -1;

            //this uses Word's Wildcard matching, which is separate from Regex. It's very limited.
            //string pattern = "(From: ?*^11)(Sent: ?*^11)(To: ?*^11)(Subject: ?*)^13";

            Word.Paragraphs paragraphs = mailInspector.Paragraphs;

            
            string rExp = @"(From: [^\n\r\v]+[\n\r\v])" +
                @"(Sent: [^\n\r\v]+[\n\r\v])" +
                @"(To: [^\n\r\v]+[\n\r\v])" +
                @"(Cc: [^\n\r\v]+[\n\r\v])?" +
                @"(Subject: [^\n\r\v]*[\n\r\v])";

            // The message property info should all be within the same paragraph.
            // Need to search paragraph by paragraph to get the specific Range. The Range.Text index would not work
            foreach(Word.Paragraph p in paragraphs)
            {
                Match match = Regex.Match(p.Range.Text, rExp);
                if (match.Success)
                {
                    startRange = p.Range.Start;
                    length = p.Range.End - startRange;
                    break;
                }
            }

            int[] propertyIndices = { startRange, length };
            return propertyIndices;
        }

        private void ExtractMessage(int endRange)
        {
            if (endRange == -1)
            {
                // no other messages in the chain, set to what's left of the message
                mailInspector.Content.Copy();
            }
            else
            {
                // copy the latest message from the chain, and remove from the content.
                mailInspector.Range(0, endRange).Copy();
                mailInspector.Range(0, endRange).Delete();
            }
        }

        private string[] ParseNextMessageProperties(int endRange)
        {
            /*
             * Using this format to parse:
             * From: X\v
             * Sent: X\v
             * To: X\v
             * Cc: X - this is optional\v
             * Subject: X\r\r
             */

            Word.Range range = mailInspector.Range(0, endRange);
            string[] split = range.Text.Split('\v');

            mailInspector.Range(0, endRange).Delete();

            string[] prop = new string[3];
            prop[0] = split[0].Remove(0, 6).TrimEnd(' '); // Remove "From: "
            prop[1] = split[1].Remove(0, 6).TrimEnd(' '); // Remove "Sent: "
            prop[2] = split[2].Remove(0, 4).TrimEnd(' '); // Remove "To: "
            if (split.Length == 5)
                prop[2] += "; " + split[3].Remove(0, 4).TrimEnd(' '); // Remove "Cc: "

            return prop;
        }
        #endregion

        #region Find Rows
        private int FindRow(string sub, DateTime dt)
        {
            int row = 0;
            Word.Table oTbl = oDoc.Tables[1];
            Word.Range rng = oTbl.Range; // Assuming there is only one table in the project notes
            object findSub = "[Subject: " + sub + "]"; // Using the new note format "[Subject: X]"

            // If subject is found in the table, then find the row.
            // Else, we can assume we can append it to the latest row
            if (rng.Find.Execute(ref findSub))
            {
                oTbl.Columns[1].Select();
                // Search all rows, from the bottom up (most recent), for any with the current mail subject

                bool gonePastThread = false;
                for (int i = oTbl.Rows.Count; i > 0; i--)
                {
                    Word.Range rowRng = oTbl.Rows[i].Range;
                    rowRng.Find.ClearFormatting();

                    if (rowRng.Sentences.Count >= 2)
                    {
                        Word.Range subjectRng = rowRng.Sentences[1];
                        Word.Range timeRng = rowRng.Sentences[2];

                        char[] trim = { '\n', '\r', ' ' };

                        if (subjectRng.Text.TrimEnd(trim).CompareTo((string)findSub) == 0)
                        {
                            gonePastThread = true;
                            int result = CompareDates(timeRng.Text.TrimEnd(trim), dt);
                            if (result < 0)
                            { //The current row has an earlier timestamp
                                row = i;
                                break;
                            }
                            else if (result > 0)
                            {
                                row = i - 1;
                                if (row == 0)
                                    row++;
                            }
                            else
                            { //Usually means the current notes are already intaken.
                                row = -1;
                            }
                        }
                        else if (gonePastThread) // If past the rows with the current subject header, break out of loof
                            break;
                    }
                    
                }
            }

            return row;
        }

        private int CompareDates(string t, DateTime dt)
        {

            string pattern = "MM-dd-yy h:mmtt"; // Using the new note format
            DateTime parsedDate = DateTime.ParseExact(t, pattern, null, System.Globalization.DateTimeStyles.AssumeLocal);

            return DateTime.Compare(parsedDate, dt);
        }

        #endregion

        #region Insert in Doc
        //TODO test to append formatted content at correct spot
        //TODO test ALL messages get included
        //TODO test to ensure embedded images and links get included in project notes
        //TODO test no inserting empty rows
        //TODO see if you can keep styling from Outlook message.
        private void InsertInDoc(string sub, string send, string rec, string t, int row)
        {
            Word.Table oTbl = oDoc.Tables[1];

            bool addToEnd = (row == 0);
            int rowCount = oTbl.Rows.Count;

            if (addToEnd) // row == 0 means the email subject wasn't found in the Project notes
            {
                row = GetEndRow(oTbl);
            }

            if (row > rowCount) // add a row to the very end
                AddRow(oTbl, oTbl.Rows[rowCount]);
            else if (!addToEnd)
            {   
                // add a row somewhere in the middle
                AddRow(oTbl, oTbl.Rows[row]);
                row += 1;
            }

            Word.Range tblRange = oTbl.Cell(row, 1).Range;

            tblRange.Paste();
            Clipboard.Clear();

            tblRange.InsertBefore(
                "[Subject: " + sub + "]\n" + //subject
                t + "\n"); //time

            if (!attachment.Equals(""))
                tblRange.InsertAfter("\n(Attachment:" + attachment + ")"); //attachments

            
            
            oTbl.Cell(row, 2).Range.Text = send + " to " + rec; //sender to receiver

        }

        // Find a row at the end of the table to append the contents to
        private int GetEndRow(Word.Table oTbl)
        {
            int rowIndex = oTbl.Rows.Count;
            int row;

            // If the last row has content in it, set the row to the next row after it
            if (oTbl.Rows[rowIndex].Range.Text.Length > 0)
                row = rowIndex + 1;
            else
            {
                // the project note documents often has empty rows. Find the earliest empty row to insert into.
                while (oTbl.Rows[rowIndex].Range.Text.Length > 0)
                {
                    rowIndex--;
                }
                row = rowIndex + 1;
            }
            return row;
        }

        // Add a row after the given row reference
        private void AddRow(Word.Table oTbl, object rowRef)
        {
            oTbl.Rows.Add(ref rowRef);
        }

        #endregion
    }
}
