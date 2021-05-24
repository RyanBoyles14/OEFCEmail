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
        // Mail properties
        private readonly string filename;
        private readonly string subject;
        private readonly string sender;
        private readonly string receiver;
        private readonly string time;
        private string attachment;

        // Document objects
        private readonly Word._Application oWord;
        private readonly Word._Document oDoc;

        // Ranges
        private int mailStartRange;
        private Word.Range mailRange;
        private Word.Range finalRange;

        // missing reference
        private object missing = Type.Missing;

        public EmailSaver(string filename, string subject, string sender, string receiver, string time, 
                            string attachment, Outlook.MailItem item)
        {
            this.filename = filename;
            this.subject = subject;
            this.sender = sender;
            this.receiver = receiver;
            this.time = time;
            this.attachment = attachment;

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

        /*
         * Append all the formatted text to the end of the project notes document
         * Requires saving the Outlook email as a temporary Document in order to preserve the formatting
         * It needs to be a separate document in order to insert the temp document into the Notes document
         * Inserting the file is the best alternative to copy/paste, which the user could accidentally use mid-function
         * After inserting the file, delete the temporary file
         */
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

        // Used to quit without saving, i.e. when the Outlook add-in encounters an error.
        public void SuspendProcess()
        {
            // In the case the document is closed when trying to close it
            try {
                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                oWord.Quit(ref saveChanges);
            } catch(Exception exc)
            {
                MessageBox.Show(exc + "\nError closing document.");
            }
            
        }

        public void Save()
        {
            bool success = true;
            if (oDoc != null)
            {
                if (!oDoc.ReadOnly) // user can still open the file, but the program cannot save to it
                {
                    bool hasMoreMessages = true;
                    bool topMessage = true;
                    int lastSearchedParagraph = 0;

                    string sub = TrimSubject(subject);
                    string send = sender;
                    string rec = receiver;
                    DateTime dt = ParseTime(time, false); // time for the top message

                    while (hasMoreMessages)
                    {
                        /* 
                         * index 0: beginning index of the next message
                         * index 1: length of the email properties to parse through
                         * index 2: the last paragraph searched for in the range.
                         */
                        int[] propRanges = GetSegmentInfo(lastSearchedParagraph);
                        int propStart = propRanges[0];
                        int propLength = propRanges[1];
                        lastSearchedParagraph = propRanges[2];

                        if (propStart == -1)
                        {
                            hasMoreMessages = false;
                            propStart = mailRange.End;
                        }

                        int row = FindRow(sub, dt, send, rec, propStart);
                        if (row == -1)
                        {
                            if (topMessage)
                            {
                                MessageBox.Show("Current message may have already been saved. Suspending process...");
                                success = false;
                            }
                            break;
                        }

                        // write to Word Doc
                        InsertInDoc(sub, send, rec, dt.ToString("MM-dd-yy h:mmtt"), row, propStart, topMessage);

                        // prepare for next cycle by getting the next message's properties.
                        if (hasMoreMessages)
                        {
                            topMessage = false;
                            string[] prop = ParseNextMessageProperties(propLength);
                            send = prop[0];
                            dt = ParseTime(prop[1], true);
                            rec = prop[2];
                            attachment = "";
                        }
                    } // endwhile (hasMoreMessages)

                    mailRange.Delete();

                    if (success)
                    {
                        oDoc.ActiveWindow.ScrollIntoView(finalRange, false);
                        oWord.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMaximize;
                        oDoc.ActiveWindow.Visible = true;
                    }
                    else
                    {
                        oWord.Quit();
                    }

                } // endif (!oDoc.ReadOnly)
            }
            else
            {
                MessageBox.Show("Word Doc couldn't open. Suspending Process...");
                SuspendProcess();
            }
        }

        #region Trim/Parse

        // Using Regex, find any messages in a chain, either forwarded or replied messages
        private int[] GetSegmentInfo(int lastSearchedParagraph)
        {
            int propStart = -1;
            int propLength = -1;

            /*
             * Using this format to find other messages in a chain
             * From: X
             * Sent: X
             * To: X
             * Cc: X - this is optional
             * Subject: X
             */
            string rExp = @"(From: [^\n\r\v]+[\n\r\v])" +
                @"(Sent: [^\n\r\v]+[\n\r\v])" +
                @"(To: [^\n\r\v]+[\n\r\v])?" +
                @"(Cc: [^\n\r\v]+[\n\r\v])?" +
                @"(Subject: [^\n\r\v]*[\n\r\v])";

            Word.Paragraphs paragraphs = mailRange.Paragraphs;

            // The message property info should all be within the same paragraph.
            // Need to search paragraph by paragraph to get the specific Range of the paragraph. Using Range.Text index would not work
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

            // Format from emails can sometimes include seconds, though it seems rare. Try parsing again if the first parse didn't work
            if(!parsed && moreMsg)
            {
                pattern = "dddd, MMMM d, yyyy h:mm:ss tt";
                parsed = DateTime.TryParseExact(t, pattern, null, System.Globalization.DateTimeStyles.AssumeLocal, out parsedDate);
            }
            
            if (!parsed) // Throw an exception if parsing the second time still doesn't work
            {
                throw new Exception("Error parsing message's sent time\n");
            }

            return parsedDate;
        }

        private string[] ParseNextMessageProperties(int length)
        {
            // Return a duplicate of the mailRange with the start and end ranges set to the global mailStartRange and a given end.
            // Must be done locally! Cannot be moved to another function
            Word.Range range = mailRange.Duplicate;
            range.Start = mailStartRange;
            range.End = mailStartRange += length;

            string[] split = range.Text.Split('\v');

            string[] prop = new string[3];

            /*
             * Using this format:
             * 
             * From: X\v
             * Sent: X\v
             * To: X\v
             * Cc: X - this is optional\v
             * Subject: X\r\r
             */
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
        private int FindRow(string sub, DateTime dt, string send, string rec, int propStart)
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
                    Word.Range contentRow = oTbl.Cell(i,1).Range;
                    contentRow.Find.ClearFormatting();

                    if (contentRow.Sentences.Count >= 2)
                    {
                        Word.Range subjectRng = contentRow.Sentences[1];
                        Word.Range timeRng = contentRow.Sentences[2];

                        char[] trim = { '\n', '\r', '\a', ' ' };

                        if (subjectRng.Text.TrimEnd(trim).CompareTo((string)findSub) == 0)
                        {
                            gonePastThread = true;

                            // the top-most message will have seconds in the timestamp
                            // Need to remove them for an equal comparison.
                            int result = CompareDates(timeRng.Text.TrimEnd(trim), dt.AddSeconds(-dt.Second)); 
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
                            {
                                // the time stamps are the same. Since forwarded messages don't include seconds,
                                // two separate messages can have the same time stamps.

                                // Return a duplicate of the mailRange with the start and end ranges set to the global mailStartRange and a given end.
                                // Must be done locally! Cannot be moved to another function
                                Word.Range mail = mailRange.Duplicate;
                                mail.Start = mailStartRange; 
                                mail.End = propStart;

                                // There can be variation between the characters surrounding an email address. Trim it, then compare
                                string senderText = oTbl.Cell(i, 2).Range.Text;
                                senderText = senderText.Replace('(', '<').Replace(')', '>').Trim(trim);
                                send = send.Replace('(', '<').Replace(')', '>').Trim(trim);
                                rec = rec.Replace('(', '<').Replace(')', '>').Trim(trim);

                                // Check if message has the same senders and receivers
                                if (senderText.Equals(send + " to " + rec))
                                {
                                    string m = "[Subject: " + sub + "]\r" + dt.ToString("MM-dd-yy h:mmtt") + "\r" + mail.Text;

                                    DialogResult dr = MessageBox.Show(
                                        "Are the contents of the two messages below the same?\r" +
                                        "If yes, the message found in the mail chain will not be saved.\r" +
                                        "-------------Message in Project Notes:-------------\r"
                                        + contentRow.Text.Trim(trim) +
                                        "\r\r------------------Message in Mail:-----------------\r"
                                        + m, "Compare Messages", MessageBoxButtons.YesNoCancel
                                        );

                                    switch (dr)
                                    {
                                        case DialogResult.Yes:
                                            row = -1; //Usually means the current notes are already intaken.
                                            break;
                                        case DialogResult.No:
                                        case DialogResult.Cancel:
                                            row = i; //intake the current message
                                            break;
                                    }


                                    /* ----Old method for comparing messages. Too inconsistent if anything in Project Notes is editted.----
                                    // two return characters always shows up after the time
                                    // Any forwarded/replied messages have return characters at the start of the message
                                    //  The top-most message does not have a return character at its start, so we have to add it in
                                    string format = "\r";
                                    if (topMessage)
                                        format += "\r";

                                    // create a string that can best match the contents in the cell if they are the same message.
                                    string m = "[Subject: " + sub + "]\r" + dt.ToString("MM-dd-yy h:mmtt") + format + mail.Text;

                                    // If these aren't attachments, we need to trim the end of the message
                                    //  (since we're also trimming the end of the cell's text for comparing)
                                    //  If there are attachments, no whitespace or return characters are included, so no trimming needed.
                                    if (!attachment.Equals(""))
                                    {
                                        m += "\r(Attachment: " + attachment.Trim(trim) + ")";
                                    }
                                    else
                                    {
                                        m = m.TrimEnd(trim);
                                    }
                                        

                                    // In the case where messages with the same sender/receiver and timestamp have different content
                                    // Compare content (assumes that no editting was involved with the table cell or the Outlook message)
                                    if (contentRow.Text.Trim(trim).Equals(m))
                                    {
                                        row = -1; //Usually means the current notes are already intaken.
                                        break;
                                    }
                                    */
                                }

                                
                            }
                        }
                        else if (gonePastThread) // If past the rows with the current subject header, break out of loop
                            break;
                    } // if (contentRow.Sentences.Count >= 2)

                } // for (int i = oTbl.Rows.Count; i > 0; i--)
            } // if (rng.Find.Execute(ref findSub))

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
        private void InsertInDoc(string sub, string send, string rec, string t, int row, int endRange, bool topMessage)
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
            }
            else if(!addToEnd)
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

            Word.Range tblRange = finalRange = oTbl.Cell(row, 1).Range;

            // Return a duplicate of the mailRange with the start and end ranges set to the global mailStartRange and a given end.
            // Must be done locally! Cannot be moved to another function
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

            // The top-most message doesn't have any empty newlines before the start
            // Insert two after the time to match the formatting of forwarded/replied messages.
            string format = "\n";
            if (topMessage)
            {
                format += "\n";
            }

            tblRange.InsertBefore(
            "[Subject: " + sub + "]\n" + t + format);

            char[] trim = { '\n', '\r', '\a', ' ' };

            if (!attachment.Equals(""))
                tblRange.InsertAfter("\n(Attachment: " + attachment.Trim(trim) + ")"); //attachments

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
