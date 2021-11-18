using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using JR.Utils.GUI.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace OEFCemail
{
    class EmailSaver
    {
        public bool Initialized = true;

        // Mail properties
        private readonly string _filename;
        private readonly string _subject;
        private readonly string _sender;
        private readonly string _recipient;
        private readonly string _time;
        private string _attachment;

        // Document objects
        private readonly Word.Application oWord;
        private readonly Word.Document oDoc;

        // Ranges
        private int mailStartRange;
        private Word.Range mailRange;
        private Word.Range finalRange;

        // missing reference
        private object missing = Type.Missing;
        
        // common newline/whitespace characters to trim
        private readonly char[] trimChars = { '\r', '\a', '\n', '\v', ' ' };

        public EmailSaver(string filename, string subject, string sender, string receiver, string time,
                            string attachment)
        {
            _filename = filename;
            _subject = subject;
            _sender = sender;
            _recipient = receiver;
            _time = time;
            _attachment = attachment;

            oWord = new Word.Application();

            try
            {
                oDoc = oWord.Documents.Open(filename, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);

                if (oDoc == null)
                {
                    FlexibleMessageBox.Show("Error Opening Word Document. Terminating Processing...");
                    Initialized = false;
                }
                else if (oDoc.ReadOnly)
                {
                    FlexibleMessageBox.Show("Word Document is Read-Only. Terminating Processing...");
                    Initialized = false;
                }
            }
            catch (Exception exc)
            {
                oWord.Quit();

                FlexibleMessageBox.Show("Error Opening Word Doc. Terminating Processing...");
                ErrorLog log = new ErrorLog();
                log.WriteErrorLog(exc.ToString());
                Initialized = false;
            }

        }

        #region Add To Document
        public async Task SaveAsync(Outlook.MailItem item)
        {
            AppendToDoc(item);
            SaveToDoc();

            await Task.Delay(0);
        }
        
        // Append all the formatted text to the end of the project notes document
        // Requires saving the Outlook email as a temporary Document in order to preserve the formatting
        // It needs to be a separate document in order to insert the temp document into the Notes document
        // Inserting the file is the best alternative to copy/paste, which the user could accidentally use mid-function
        // After inserting the file, delete the temporary file
        public void AppendToDoc(Outlook.MailItem item)
        {
            object path = System.IO.Path.GetDirectoryName(_filename) + "\\(temporary).doc";
            object format = Word.WdSaveFormat.wdFormatDocument;
            Word.Document mailInspector = item.GetInspector.WordEditor as Word.Document;

            mailInspector.SaveAs2(ref path, ref format);

            mailStartRange = oDoc.Content.End - 1;
            mailRange = oDoc.Range(mailStartRange, ref missing);
            mailRange.InsertFile((string)path, ref missing, ref missing, ref missing, ref missing);

            System.IO.File.Delete((string)path);

        }

        // Used to quit without saving, i.e. when the Outlook add-in encounters an error.
        public void TerminateProcess()
        {
            // In the case an error occurs when trying to close Word
            try
            {
                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                oWord.Quit(ref saveChanges);
            }
            catch (Exception exc)
            {
                FlexibleMessageBox.Show(exc + "\nError closing Word.");

                ErrorLog log = new ErrorLog();
                log.WriteErrorLog(exc.ToString());
            }

        }

        public void SaveToDoc()
        {
            bool success = true;
            bool haveMoreMessages = true;
            bool isTopMessage = true;
            int lastSearchedParagraph = 0;
            int row = 0;

            string sub = TrimSubject(_subject);
            string send = _sender;
            string rec = _recipient;
            DateTime dt = ParseTime(_time, true); // DateTime for the topmost message

            while (haveMoreMessages)
            {
                // if there are any other messages in the email chain, get the properities of its range.
                // currentMsgEnd = beginning range of the next message, also used as the end of the current message's range
                // nextMsgLength = length of the next messages segment's 
                (int currentMsgEnd, int nextMsgLength, int paragraph) = GetNextMsgProperties(lastSearchedParagraph);
                lastSearchedParagraph = paragraph; // the last paragraph searched for in the mail's range.

                // if no other messages
                if (currentMsgEnd == -1)
                {
                    haveMoreMessages = false;
                    currentMsgEnd = mailRange.End;
                }

                row = FindRow(sub, dt, currentMsgEnd, row);

                // row = -1 means the current msg has already been saved into the table.
                if (row == -1)
                {
                    if (isTopMessage)
                    {
                        FlexibleMessageBox.Show("Current email thread may have already been saved. Terminating process...");
                        success = false;
                    }
                    break;
                }

                // write to Word Doc
                InsertInRow(sub, send, rec, dt.ToString("MM-dd-yy h:mmtt"), row, currentMsgEnd, isTopMessage);

                // prepare for next cycle by getting the next message's properties.
                if (haveMoreMessages)
                {
                    isTopMessage = false;
                    (string s1, string s2, string s3) = ParseNextMsgInfo(nextMsgLength);
                    send = s1;
                    dt = ParseTime(s2, false);
                    rec = s3;
                    _attachment = "";
                }
            } // endwhile (hasMoreMessages)

            mailRange.Delete();

            Quit(success);
        }

        private void Quit(bool success)
        {
            
             // if the email saved correctly, scroll the view of the document to the row with the topmost message
             // Bring the document into view for the user.
            if (success)
            {
                oDoc.ActiveWindow.ScrollIntoView(finalRange, false);
                oDoc.ActiveWindow.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMaximize;
                oDoc.ActiveWindow.Visible = true;
                oDoc.Activate();
                oWord.Activate();
            }
            else
            {
                oWord.Quit();
            }
        }

        #endregion

        #region Trim/Parse

        // Get the base subject header w/o Forward or Reply prefixes
        private static string TrimSubject(string sub)
        {
            // Typical abbreviations that are prepended to subject lines are: "FW", "Fwd", or "RE". Search for any variation of those.
            // Configured in case other prefixes need to be added.
            List<String> abbr = new List<String> { "RE:", "FW:", "FWD:" };

            // using Linq: return the first string "s" in abbr where sub.ToUpper() starts with "s". If none exist, returns null.
            // https://stackoverflow.com/questions/12296089/how-to-check-if-a-string-contains-any-element-of-a-liststring
            String str = abbr.FirstOrDefault(s => sub.ToUpper().StartsWith(s));

            // In some instances, the subject line can include multiple of these abbreviations
            while (str != null)
            {
                sub = sub.Remove(0, str.Length);

                // Trim any leftover spaces that may remain before the subject line
                sub = sub.TrimStart(' ');

                str = abbr.FirstOrDefault(s => sub.ToUpper().StartsWith(s));
            }

            return sub;
        }

        
        // Using Regex, find any messages in a chain, either forwarded or replied messages
        // Return a 3-integer Tuple with the properties of the next forwarded/replied message within the email chain
        // int1: beginning range of the next message, also used as the end of the current message's range
        // int2: length of the next message
        // int3: the last paragraph searched for in the mail's range.
         
        private (int, int, int) GetNextMsgProperties(int lastSearchedParagraph)
        {
            int msgStart = -1;
            int msgLength = -1;


            // Using this format to find other messages in a chain
            // From: X
            // Sent: X
            // To: X
            // Cc: X - this is optional
            // Subject: X
            string rExp = @"(From: [^\n\r\v]+[\n\r\v])" +
                @"(Sent: [^\n\r\v]+[\n\r\v])" +
                @"(To: [^\n\r\v]+[\n\r\v])?" +
                @"(Cc: [^\n\r\v]+[\n\r\v])?" +
                @"(Subject: [^\n\r\v]*[\n\r\v])";

            Word.Paragraphs paragraphs = mailRange.Paragraphs;

            // The message property info should all be within the same paragraph.
            // Need to search paragraph by paragraph to get the specific Range of the paragraph. Using Range.Text index would not work
            // Adding 1 to lastSearchedParagraph, since Word object indices are 1-based, not 0-based
            for (int i = lastSearchedParagraph + 1; i <= paragraphs.Count; i++)
            {
                Word.Paragraph p = paragraphs[i];
                Match match = Regex.Match(p.Range.Text, rExp);
                if (match.Success)
                {
                    msgStart = p.Range.Start;
                    msgLength = p.Range.End - msgStart;
                    lastSearchedParagraph = i;
                    break;
                }
            }

            return (msgStart, msgLength, lastSearchedParagraph);
        }

        // parse the time based on the typical formats from Outlook emails
        // https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings
        private DateTime ParseTime(string time, bool isFirstMsg)
        {
            string pattern;
            if (isFirstMsg)
            {
                // Format from for the timestamp for the overall email item (the topmost message): "x/xx/20xx x:xx:xx xM"
                pattern = "M/d/yyyy h:mm:ss tt";
            }
            else
            {
                // Using the Format from emails in the chain: "Day, Month d, yyyy h:mm xM"
                pattern = "dddd, MMMM d, yyyy h:mm tt";
            }

            bool parsed = DateTime.TryParseExact(time, pattern, null, System.Globalization.DateTimeStyles.AssumeLocal, out DateTime parsedDate);

            // Format from emails in the chain can sometimes include seconds, though it seems rare. Try parsing again if the first parse didn't work
            if (!parsed && !isFirstMsg)
            {
                pattern = "dddd, MMMM d, yyyy h:mm:ss tt";
                parsed = DateTime.TryParseExact(time, pattern, null, System.Globalization.DateTimeStyles.AssumeLocal, out parsedDate);
            }

            if (!parsed) // Throw an exception if parsing the second time still doesn't work
            {
                throw new Exception("Error parsing message's sent time\r" +
                    "Error occurred at text: \"" + time + "\"");
            }

            return parsedDate;
        }

        private (string, string, string) ParseNextMsgInfo(int length)
        {
            Word.Range range = DuplicateMailRange(mailStartRange + length);
            mailStartRange += length;

            string[] split = range.Text.Split('\v');

            string sender;
            string time;
            string recipient;


            //Most emails use this format:
            //From: X\v
            //Sent: X\v
            //To: X\v
            //Cc: X\v - this is optional
            // Subject: X\r\r
            try
            {
                sender = split[0].Remove(0, 6).TrimEnd(' '); // Remove "From: "
                time = split[1].Remove(0, 6).TrimEnd(' '); // Remove "Sent: "
                recipient = split[2].Remove(0, 4).TrimEnd(' '); // Remove "To: "
            }
            catch (Exception exc)
            {
                throw new Exception("Error parsing messages in the chain.\r" +
                    "Error found at text:\r" +
                    "\"" + split[0] + "\r" + split[1] + "\r" + split[2] + "\"", exc);
            }

            // if it includes CC'd recipients, add to the list of recipients
            if (split.Length == 5)
                recipient += "; " + split[3].Remove(0, 4).TrimEnd(' '); // Remove "Cc: "

            return (sender, time, recipient);
        }
        #endregion

        #region Duplicate Range

        // Return a duplicate of the mailRange with the start and end ranges set to the global mailStartRange and a given end.
        private Word.Range DuplicateMailRange(int end)
        {
            Word.Range mail = mailRange.Duplicate;
            mail.Start = mailStartRange;
            mail.End = end;

            return mail;
        }

        #endregion

        #region Compare

        private DialogResult CompareMessages(Word.Range contentRow, String sub, DateTime dt, int propStart)
        {
            Word.Range mail = DuplicateMailRange(propStart);

            string m = "[Subject: " + sub + "]\r" + dt.ToString("MM-dd-yy h:mmtt") + "\r" + mail.Text;

            return FlexibleMessageBox.Show(
                "Are the contents of the two messages below the same?\r" +
                "If yes, the message found in the mail chain will not be saved.\r\r" +
                "-------------Message in Project Notes:-------------\r\r"
                + contentRow.Text.Trim(trimChars) +
                "\r\r\r------------------Message in Mail:-----------------\r\r"
                + m, "Compare Messages", MessageBoxButtons.YesNoCancel
            );
        }

        private int CompareDates(string t, DateTime dt)
        {
            string pattern = "MM-dd-yy h:mmtt"; // Using the new note format
            DateTime parsedDate = DateTime.ParseExact(t, pattern, null, System.Globalization.DateTimeStyles.AssumeLocal);

            return DateTime.Compare(parsedDate, dt);
        }

        #endregion

        #region Find Rows 
        private int FindRow(string sub, DateTime dt, int propStart, int row)
        {

            Word.Table oTbl = oDoc.Tables[1];
            Word.Range rng = oTbl.Range; // Assuming there is only one table in the project notes
            object findSub = "[Subject: " + sub + "]"; // Using the new note format "[Subject: X]"

            // row = 0 or row = -1 means the prev message was added to the end of the table
            // or means the current message is the first message in the thread
            int i = (row == 0 || row == -1) ? oTbl.Rows.Count : row;

            // If mail subject is found in the table, then find the row.
            // Else, we can assume we can append it to the latest row
            if (rng.Find.Execute(ref findSub))
            {
                oTbl.Columns[1].Select();

                // Search for the current mail subject in all rows, starting from the bottom (most recent)
                // or from the previously saved message
                bool foundSubjectHeader = false;
                for (; i > 0; i--)
                {
                    Word.Range contentRow = oTbl.Cell(i, 1).Range;
                    contentRow.Find.ClearFormatting();

                    if (contentRow.Sentences.Count >= 2)
                    {
                        Word.Range subjectRng = contentRow.Sentences[1];
                        Word.Range timeRng = contentRow.Sentences[2];

                        if (subjectRng.Text.TrimEnd(trimChars).CompareTo((string)findSub) == 0)
                        {
                            foundSubjectHeader = true;

                            // the latest message (message received) will have seconds in the timestamp
                            // Need to remove the seconds for an equal comparison with other messages.
                            int result = CompareDates(timeRng.Text.TrimEnd(trimChars), dt.AddSeconds(-dt.Second));
                            if (result < 0)
                            {   //The current row has an earlier timestamp than the message we want to intake
                                //Treat the current row as our reference point for inserting the message.
                                row = i;
                                break;
                            }
                            else if (result == 0)
                            {
                                // the time stamps are the same. Since forwarded messages don't include seconds,
                                // two separate messages can have the same time stamps.
                                DialogResult dr = CompareMessages(contentRow, sub, dt, propStart);

                                switch (dr)
                                {
                                    case DialogResult.Yes:
                                        row = -1;
                                        break;
                                    case DialogResult.No:
                                    case DialogResult.Cancel:
                                        row = i; //intake the current message
                                        break;
                                }

                                if (row == -1)
                                    // don't need to search through the notes any more. Assume the rest of the email thread is already intaken
                                    break;

                            }
                        }

                        // If the thread of emails with the current subject header is found,
                        // but the current row no longer has that subject header, break out of loop
                        // This will return the current row right before the first row with the subject header
                        else if (foundSubjectHeader)
                        {
                            row = i;
                            break;
                        }
                    } // if (contentRow.Sentences.Count >= 2)

                } // for (int i = oTbl.Rows.Count; i > 0; i--)
            } // if (rng.Find.Execute(ref findSub))

            return row;
        }
        #endregion

        #region Insert into Row
        private void InsertInRow(string sub, string send, string rec, string t, int row, int endRange, bool topMessage)
        {
            Word.Table oTbl = oDoc.Tables[1];

            int rowCount = oTbl.Rows.Count;
            bool addToEnd = (row == 0) || (row == rowCount);

            int startRangeBeforeCopy = mailRange.Start;
            int diff;
            int length = endRange - mailStartRange;

            if (addToEnd)
            {
                // int row represents the exact row to insert the message into
                row = GetLastRow(oTbl, rowCount);

                // row > rowCount only when the table has no empty rows (we need to add a new empty row)
                if (row > rowCount)
                {
                     // For some reason, the program can't simply append an empty row to the end of the table.
                     // Instead, it will insert an empty row right before the last row.
                     // This is a work around. Insert the empty row right before the last row,
                     // then copy the last row into the new empty row.
                     // This allows the last row to be overwritten.
                    AddNewRow(oTbl, oTbl.Rows[rowCount]);

                    object oChar = Word.WdUnits.wdCharacter; //represents a single character in Word.

                    for (int i = 1; i < oTbl.Rows[rowCount].Cells.Count; i++)
                    {
                        Word.Range copyFrom = oTbl.Rows[row].Cells[i].Range;

                        //moving the end allows the program to only copy the contents of the cell, and prevents copying the end of the Cell
                        copyFrom.MoveEnd(ref oChar, -1);

                        oTbl.Rows[rowCount].Cells[1].Range.FormattedText = copyFrom.FormattedText;
                    }

                }
            }
            else
            {
                // add a row somewhere within the table
                // increment row to the row we want to insert into
                AddNewRow(oTbl, oTbl.Rows[++row]);
            }

            // Adding a new row may cause the ranges for the mail messages to shift. Update as needed
            if (startRangeBeforeCopy != mailRange.Start)
            {
                diff = mailRange.Start - startRangeBeforeCopy;
                mailStartRange += diff;
                endRange += diff;
                startRangeBeforeCopy = mailRange.Start;
            }

            Word.Range tblRange = finalRange = oTbl.Cell(row, 1).Range;

            Word.Range range = DuplicateMailRange(endRange);

            // Copy from the message's range into the table row's range.
            try
            {
                //TODO: change text-font to that of our template?
                tblRange.FormattedText = range.FormattedText;
            }
            catch (Exception exc)
            {
                throw new Exception("Error copying text over.\r" +
                    "Error occurred at text: \"" + range.Text + "\"", exc);
            }

            // The top-most message doesn't have any empty newlines before the start
            // For consistent formatting across all messages,
            // Insert two newlines after the time to match the formatting of forwarded/replied messages.
            // TODO check if necessary. Seems to mess up
            string format = "\n";
            if (topMessage)
            {
                format += "\n";
            }

            tblRange.InsertBefore(
            "[Subject: " + sub + "]\n" + t + format);

            if (!_attachment.Equals(""))
                tblRange.InsertAfter("\n(Attachment: " + _attachment.Trim(trimChars) + ")"); //attachments

            oTbl.Cell(row, 2).Range.Text = send + " to " + rec; //sender to receiver

            // Copying the mail message into the row may cause the ranges for the mail messages to shift. Update as needed
            if (startRangeBeforeCopy != mailRange.Start)
            {
                diff = mailRange.Start - startRangeBeforeCopy;
                mailStartRange += diff + length;
            }
        }

        // Add a row after the given row reference
        private void AddNewRow(Word.Table oTbl, object rowRef)
        {
            oTbl.Rows.Add(ref rowRef);
        }

        // Find a row at the end of the table to append the contents to
        private int GetLastRow(Word.Table oTbl, int rowCount)
        {
            int row;

            // If the very last row of the table has content in it
            if (oTbl.Rows[rowCount].Range.Text.Trim(trimChars).Length > 0)
                row = rowCount + 1;
            else // the project note documents often has empty rows. Find the earliest empty row to insert into.
            {
                while (oTbl.Rows[rowCount].Range.Text.Trim(trimChars).Length == 0)
                {
                    rowCount--;
                    if (rowCount == 0)
                        break;
                }
                row = rowCount + 1;
            }
            return row;
        }

        #endregion
    }

}
