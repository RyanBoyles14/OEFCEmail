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
        private int mailStartRange;
        private Word.Range mailRange;
        private object missing = Type.Missing;

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
                oDoc = oWord.Documents.Open(filename, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);
                AppendDoc(item);
            }
            catch (Exception exc)
            {
                if (exc is IOException)
                    MessageBox.Show(exc + "\nError Opening Word Doc. Check that it is not already open");
            }
        }

        // Append all the formatted text to the end of the project notes document
        private void AppendDoc(Outlook.MailItem item)
        {
            object path = System.IO.Path.GetDirectoryName(filename) + "\\(temporary).doc";
            object format = Word.WdSaveFormat.wdFormatDocument;
            Word.Document mailInspector = item.GetInspector.WordEditor as Word.Document;
            mailInspector.SaveAs2(ref path, ref format);

            mailStartRange = oDoc.Content.End - 1;
            mailRange = oDoc.Range(mailStartRange, ref missing);
            mailRange.InsertFile((string)path, ref missing, ref missing, ref missing, ref missing);

            System.IO.File.Delete((string)path);
        }

        // Only needed to quit without saving, i.e. errors.
        public void SuspendProcess()
        {
            try {
                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                oWord.Quit(ref saveChanges);
            } catch(Exception exc)
            {
                MessageBox.Show(exc + "\nError closing document.");
            }
            
        }

        //TODO progress bar?
        //TODO trim down whitespace/signoffs?
        public void Save()
        {
            bool success = true;
            if (oDoc != null)
            {
                if (!oDoc.ReadOnly) // user can still open the file, but the program cannot save to it
                {
                    bool hasMoreMessages = true;
                    bool firstMessage = true;
                    int lastSearchedParagraph = 0;
                    string sub = TrimSubject(subject);
                    string send = sender;
                    string rec = receiver;
                    DateTime dt = ParseTime(time, false); // time for the top message
                    while (hasMoreMessages)
                    {
                        // index 0: beginning index of the next message
                        // index 1: length of the email properties to parse through
                        // index 2: the last paragraph searched for in the range.
                        int[] propRanges = GetSegmentInfo(lastSearchedParagraph);
                        int propStart = propRanges[0];
                        int propLength = propRanges[1];
                        lastSearchedParagraph = propRanges[2];

                        if (propStart == -1)
                        {
                            hasMoreMessages = false;
                            propStart = mailRange.End;
                        }

                        int row = FindRow(sub, dt);
                        if (row == -1)
                        {
                            if (firstMessage)
                            {
                                MessageBox.Show("Current message may have already been saved. Suspending process...");
                                success = false;
                            }
                            break;
                        }

                        // write to Word Doc
                        InsertInDoc(sub, send, rec, dt.ToString("MM-dd-yy h:mmtt"), row, propStart, firstMessage);

                        // prepare for next cycle by getting the next message's properties.
                        if (hasMoreMessages)
                        {
                            firstMessage = false;
                            string[] prop = ParseNextMessageProperties(propLength);
                            send = prop[0];
                            dt = ParseTime(prop[1], true);
                            rec = prop[2];
                        }
                    }
                    if (success)
                    {
                        mailRange.Delete();
                        oDoc.ActiveWindow.Visible = true;
                    }
                    else
                        oWord.Quit();
                }
            }
            else
            {
                MessageBox.Show("Word Doc couldn't open. Suspending Process...");
                SuspendProcess();
            }
        }

        #region Trim/Parse

        /*
         * Using Regex, find other messages down the chain
         * Using this format to find other messages in a chain
         * From: X
         * Sent: X
         * To: X
         * Cc: X - this is optional
         * Subject: X
         */
        private int[] GetSegmentInfo(int lastSearchedParagraph)
        {
            int propStart = -1;
            int propLength = -1;

            string rExp = @"(From: [^\n\r\v]+[\n\r\v])" +
                @"(Sent: [^\n\r\v]+[\n\r\v])" +
                @"(To: [^\n\r\v]+[\n\r\v])?" +
                @"(Cc: [^\n\r\v]+[\n\r\v])?" +
                @"(Subject: [^\n\r\v]*[\n\r\v])";

            Word.Paragraphs paragraphs = mailRange.Paragraphs;

            // The message property info should all be within the same paragraph.
            // Need to search paragraph by paragraph to get the specific Range. The Range.Text index would not work
            for (int i = lastSearchedParagraph + 1; i <= paragraphs.Count; i++)
            {
                Word.Paragraph p = paragraphs[i];
                Match match = Regex.Match(p.Range.Text, rExp);
                if (match.Success)
                {
                    propStart = p.Range.Start;
                    propLength = p.Range.End - propStart;
                    lastSearchedParagraph = i;
                    break;
                }
            }

            int[] propertyIndices = { propStart, propLength, lastSearchedParagraph};
            return propertyIndices;
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
        private DateTime ParseTime(string t, bool moreMsg)
        {
            string pattern;
            if (moreMsg)
            {
                // Using the Format from emails in the chain: "Day, Month d, yyyy h:mm xM"
                // https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings
                pattern = "dddd, MMMM d, yyyy h:mm tt";
            }
            else
            {
                // Format from textBoxTime.Text: "x/xx/20xx x:xx:xx xM"
                pattern = "M/d/yyyy h:mm:ss tt";
            }

            bool parsed = DateTime.TryParseExact(t, pattern, null, System.Globalization.DateTimeStyles.AssumeLocal, out DateTime parsedDate);

            // Format from emails can sometimes include seconds, though it seems rare. Try parsing again.
            if(!parsed && moreMsg)
            {
                pattern = "dddd, MMMM d, yyyy h:mm:ss tt";
                parsed = DateTime.TryParseExact(t, pattern, null, System.Globalization.DateTimeStyles.AssumeLocal, out parsedDate);
            }
            
            if (!parsed) // Throw an exception if parsing the time still doesn't work
            {
                throw new Exception("Error parsing message's sent time\n");
            }

            return parsedDate;
        }

        private string[] ParseNextMessageProperties(int length)
        {
            /*
             * Using this format to parse:
             * From: X\v
             * Sent: X\v
             * To: X\v
             * Cc: X - this is optional\v
             * Subject: X\r\r
             */

            Word.Range range = mailRange.Duplicate;
            range.Start = mailStartRange;
            range.End = mailStartRange += length;

            string[] split = range.Text.Split('\v');

            string[] prop = new string[3];
            try
            {
                prop[0] = split[0].Remove(0, 6).TrimEnd(' '); // Remove "From: "
                prop[1] = split[1].Remove(0, 6).TrimEnd(' '); // Remove "Sent: "
                prop[2] = split[2].Remove(0, 4).TrimEnd(' '); // Remove "To: "
            } catch
            {
                throw new Exception("Error parsing messages in the chain.\n");
            }
            if (split.Length == 5)
                prop[2] += "; " + split[3].Remove(0, 4).TrimEnd(' '); // Remove "Cc: "

            return prop;
        }
        #endregion

        #region Find Rows
        //TODO test saving to files some of the email thread already saved.
        private int FindRow(string sub, DateTime dt)
        {
            int row = -2;
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

                            //TODO: Include a check in case two DateTimes are the same, but the content hasn't been included.
                            int result = CompareDates(timeRng.Text.TrimEnd(trim), dt);
                            if (result < 0)
                            { //The current row has an earlier timestamp
                                row = i;
                                break;
                            }
                            else if (result > 0)
                            {
                                row = i - 1;
                            }
                            else
                            { //Usually means the current notes are already intaken.
                                row = -1;
                                break;
                            }
                        }
                        else if (gonePastThread) // If past the rows with the current subject header, break out of loop
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
        //TODO test no inserting empty rows/deleting existing rows (specifically with email threads)
        private void InsertInDoc(string sub, string send, string rec, string t, int row, int endRange, bool firstMessage)
        {
            Word.Table oTbl = oDoc.Tables[1];

            bool addToEnd = (row == -2);
            int rowCount = oTbl.Rows.Count;
            int start = mailRange.Start;
            int diff;
            int length = endRange - mailStartRange;

            if (addToEnd) // the email subject wasn't found in the Project notes
            {
                // int row represents the exact row to insert the message into
                row = GetLastRow(oTbl, rowCount);
            }

            if (row > rowCount)
            {
                // Inserting rows isn't convenient. It will insert it before the very last row.
                // This is a work around. Add a row to the very end by copying the last row's contents into the row before it.
                InsertRow(oTbl, oTbl.Rows[rowCount]);

                object oChar = Word.WdUnits.wdCharacter;

                Word.Range copyFrom = oTbl.Rows[row].Cells[1].Range;
                copyFrom.MoveEnd(ref oChar, -1);
                oTbl.Rows[rowCount].Cells[1].Range.FormattedText = copyFrom.FormattedText;

                copyFrom = oTbl.Rows[row].Cells[2].Range;
                copyFrom.MoveEnd(ref oChar, -1);
                oTbl.Rows[rowCount].Cells[2].Range.FormattedText = copyFrom.FormattedText;
            } else if(!addToEnd)
            {
                // add a row somewhere in the middle
                // increment row to the row we want to insert into
                InsertRow(oTbl, oTbl.Rows[++row]);
            }

            if (start != mailRange.Start) // Update the section ranges if the mailRange ranges updated
            {
                diff = mailRange.Start - start;
                mailStartRange += diff;
                endRange += diff;
                start = mailRange.Start;
            }

            Word.Range tblRange = oTbl.Cell(row, 1).Range;

            Word.Range range = mailRange.Duplicate;
            range.Start = mailStartRange;
            range.End = endRange;

            // When the projects notes table have no empty rows, it can't copy from one range to the other. This shouldn't usually happen based on our template document.
            try
            {
                tblRange.FormattedText = range.FormattedText;
            }
            catch 
            {
                throw new Exception("Error copying text over.\n");
            }

            string format = "\n";
            if (firstMessage)
                format += "\n";

            tblRange.InsertBefore(
            "[Subject: " + sub + "]\n" + t + format);

            if (!attachment.Equals(""))
                tblRange.InsertAfter("\n(Attachment:" + attachment + ")"); //attachments

            oTbl.Cell(row, 2).Range.Text = send + " to " + rec; //sender to receiver

            if (start != mailRange.Start) // Update the section ranges if the mailRange ranges updated
            {
                diff = mailRange.Start - start;
                mailStartRange += diff + length;
            }
        }

        // Find a row at the end of the table to append the contents to
        private int GetLastRow(Word.Table oTbl, int rowCount)
        {
            int row;

            char[] trim = { '\r', '\a', '\n', '\v', ' ' };
            // If the very last row has content in it, set the row to the next row after it
            if (oTbl.Rows[rowCount].Range.Text.Trim(trim).Length > 0)
                row = rowCount + 1;
            else
            {
                // the project note documents often has empty rows. Find the earliest empty row to insert into.
                while (oTbl.Rows[rowCount].Range.Text.Trim(trim).Length == 0)
                {
                    rowCount--;
                    if (rowCount == 0)
                        break;
                }
                row = rowCount + 1;
            }
            return row;
        }

        // Add a row after the given row reference
        private void InsertRow(Word.Table oTbl, object rowRef)
        {
            oTbl.Rows.Add(ref rowRef);
        }
        #endregion
    }

}
