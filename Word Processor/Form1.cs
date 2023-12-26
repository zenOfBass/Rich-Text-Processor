using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Media;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Rich_Text_Processor // this program will be a word processor based around the rich text box
{ // open namespace
    public partial class Form1 : Form
    { // open form
        public Form1() => InitializeComponent(); // build main form

        private string currentDoc; // string of current document's name
        private int linesPrinted;  // line counter for printing
        private string[] lines;    // array of strings for printing lines

        #region File Menu Methods

        private void NewToolStripMenuItem_Click(object sender, EventArgs e) // new document click event
        { // open event
            try // try to open new document
            { // open try block
                if (textBoxEditor.Modified) // if the editor's text box has been modified
                { // open main if statement
                    DialogResult answer = MessageBox.Show(                     // show user message box
                        "Save current document before creating new document?", // ask if they would like to save the current document
                        "Unsaved Document",                                    // header
                        MessageBoxButtons.YesNo,                               // ask user for yes or no input
                        MessageBoxIcon.Question);
                    if (answer == DialogResult.No)                             // if the user answers no
                    { // open nested if statement
                        currentDoc = "";                            // empty current file
                        Text = "Rich Text Processor: New Document"; // header
                        textBoxEditor.Modified = false;             // set editor to unmodified
                        textBoxEditor.Clear();                      // clear editor
                        return;
                    } // close nested if statement
                    else // the user answers yes
                    { // open nested else statement
                        SaveToolStripMenuItem_Click(this, new EventArgs()); // user save menu click event method
                        textBoxEditor.Modified = false;                     // set editor to unmodified
                        textBoxEditor.Clear();                              // clear editor
                        currentDoc = "";                                    // empty current file
                        Text = "Rich Text Processor: New Document";         // header
                        return;
                    } // close nested else
                } // close main if statement
                else // otherwise the editor has not been modified
                { // open else statement
                    currentDoc = "";                            // empty current file
                    Text = "Rich Text Processor: New Document"; // header
                    textBoxEditor.Modified = false;             // set editor to unmodified
                    textBoxEditor.Clear();                      // clear editor
                    return;
                } // close else statement
            } // close try block
            catch // opening a new document fails
            { // open catch block
                SystemSounds.Hand.Play(); // play error sound
            } // close catch block
        } // close event
        private void OpenToolStripMenuItem_Click(object sender, EventArgs e) // open document click event
        { // open event
            try // try to open a document from file
            { // open try block
                if (textBoxEditor.Modified) // if the editor has been modified
                { // open if statement
                    DialogResult answer = MessageBox.Show(                         // show user message box
                        "Save current file before opening another document?",      // ask if they would like to save the current document
                        "Unsaved Document",                                        // header
                        MessageBoxButtons.YesNo,                                   // ask user for yes or no input
                        MessageBoxIcon.Question);
                    if (answer == DialogResult.No) textBoxEditor.Modified = false; // if they answer no the editor is set to unmodified
                    else SaveToolStripMenuItem_Click(this, new EventArgs());       // otherwise, use save event method
                } // close if statement 
                OpenFile(); // use the open file method
            } // close try block
            catch // opening a document fails
            { // open catch block
                SystemSounds.Hand.Play(); // play error sound
            } // close catch block
        }
        private void OpenFile() // open file method
        { // open method
            try // try to open a file
            { // open try block
                openFileDialog.Title = "RTP - Open File";                                                        // header
                openFileDialog.DefaultExt = "rtf";                                                               // default file extension
                openFileDialog.Filter = "Rich Text Files|*.rtf|Text Files|*.txt|HTML Files|*.htm|All Files|*.*"; // files types to fiter for
                openFileDialog.FilterIndex = 1;                                                                  // filter index number set to 1
                openFileDialog.FileName = string.Empty;                                                          // file name set to empty string
                if (openFileDialog.ShowDialog() == DialogResult.OK) // if user answers ok
                { // open main if statement
                    if (openFileDialog.FileName == "") return; // if the file name is empty, return
                    string strExt = System.IO.Path.GetExtension(openFileDialog.FileName).ToUpper(); // the file extension is retrived as a string and set to upper case
                    if (strExt == ".RTF") textBoxEditor.LoadFile(openFileDialog.FileName, RichTextBoxStreamType.RichText); // if the string extension is .RTF, load the file
                    else // otherwise, the files is of a different type, so we prepare
                    { // open nested else statement
                        System.IO.StreamReader txtReader = new System.IO.StreamReader(openFileDialog.FileName); // open the stream reader
                        textBoxEditor.Text = txtReader.ReadToEnd();                                             // read the file to the editor
                        txtReader.Close();                                                                      // close stream reader
                        txtReader = null;                                                                       // set reader to null
                        textBoxEditor.SelectionStart = 0;                                                       // reset text selection
                        textBoxEditor.SelectionLength = 0;
                    } // close nested else statement
                    currentDoc = openFileDialog.FileName;                   // current file is set
                    textBoxEditor.Modified = false;                         // set editor to unmodified
                    Text = "Rich Text Processor: " + currentDoc.ToString(); // new header with file name
                } // close main if statement
                else _ = MessageBox.Show("Open File request cancelled by user.", "Cancelled"); // otherwise, the user cancels
            } // close try block
            catch // opening a file fails
            { //  open catch block
                SystemSounds.Hand.Play(); // play error sound
            } // close catch block
        } // close method
        private void SaveToolStripMenuItem_Click(object sender, EventArgs e) // save menu event 
        { // open event
            try // try to save current document
            { // open try block
                saveFileDialog.Title = "RTP - Save File";                                                        // header
                saveFileDialog.DefaultExt = "rtf";                                                               // defaut extension to save under
                saveFileDialog.Filter = "Rich Text Files|*.rtf|Text Files|*.txt|HTML Files|*.htm|All Files|*.*"; // files filter
                saveFileDialog.FilterIndex = 1;                                                                  // index set to 1
                if (saveFileDialog.ShowDialog() == DialogResult.OK) // if the user clicks ok
                { // open main if statement
                    if (saveFileDialog.FileName == "") return; // if the file name is empty, return
                    string strExt = System.IO.Path.GetExtension(saveFileDialog.FileName).ToUpper(); // get extension as string, set to upper case
                    if (strExt == ".RTF") textBoxEditor.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.RichText); // if extension is .RTF, save
                    else // otherwise it is of another file type, and we have some things to do for that case
                    { // open nested else statement
                        System.IO.StreamWriter txtWriter = new System.IO.StreamWriter(saveFileDialog.FileName); // open stream writer
                        txtWriter.Write(textBoxEditor.Text);                                                    // write the text from the editor to the file
                        txtWriter.Close();                                                                      // close the writer
                        txtWriter = null;                                                                       // set writer to null
                        textBoxEditor.SelectionStart = 0;                                                       // reset selection
                        textBoxEditor.SelectionLength = 0;
                    } // closer nested else statement
                    currentDoc = saveFileDialog.FileName;                            // current file name is set
                    textBoxEditor.Modified = false;                                  // set to unmodified
                    Text = "Rich Text Processor: " + currentDoc.ToString();          // header with current file name
                    _ = MessageBox.Show(currentDoc.ToString() + " saved.", "Saved"); // confirm results to user
                } // close main if statement
                else _ = MessageBox.Show("Save File request cancelled by user.", "Cancelled"); // otherwise, the user cancled the request
            } // close try block
            catch // saving fails
            { // open catch block
                SystemSounds.Hand.Play(); // play error sound
            } // close catch block
        } // close method
        private void ExitToolStripMenuItem_Click(object sender, EventArgs e) // exit menu click event
        { // open event
            try // try to exit the editor
            { // open try block
                if (textBoxEditor.Modified) // if the editor has been modified
                { // open main if
                    DialogResult answer = MessageBox.Show(    // message box to user
                        "Save this document before closing?", // ask if they want to save
                        "Unsaved Document",                   // header
                        MessageBoxButtons.YesNo,              // yes no parameters
                        MessageBoxIcon.Question);
                    if (answer == DialogResult.Yes) return; // if they answer yes, return
                    else // otherwise
                    { // open nested else
                        textBoxEditor.Modified = false; // set to unmodified
                        Application.Exit();             // use exit method
                    } // close nested else
                } // close main if
                else // otherwise, the editor hasn't been used
                { // open else
                    textBoxEditor.Modified = false; // set unnmodified
                    Application.Exit(); //          // use exit method
                } // close else
            } // close try block
            catch // exiting form fails
            { // open catch block
                SystemSounds.Hand.Play(); // error sound
            } // close catch block
        } // close event
        #endregion // end of file menu


        #region Edit Menu Methods

        private void SelectAllToolStripMenuItem_Click(object sender, EventArgs e) // select all menu event
        { // open event
            try // try to selct all
            { // open try block
                textBoxEditor.SelectAll(); // select all method
            } // close try block
            catch (Exception) // can't select all text
            { // open catch block
                _ = MessageBox.Show("Unable to select all document contents.", "RTP - Select all", MessageBoxButtons.OK, MessageBoxIcon.Error); // error message
            } // close catch block
        } // close event
        private void CopyToolStripMenuItem_Click(object sender, EventArgs e) // copy menu event
        { // open event
            try // try to copy text
            { // open try block
                textBoxEditor.Copy(); // use copy method
            } // close try block
            catch (Exception) // can't copy text
            { // open catch block
                _ = MessageBox.Show("Unable to copy document contents.", "RTP - Copy", MessageBoxButtons.OK, MessageBoxIcon.Error); // error message
            } //  close catch block
        } //  close event
        private void CutToolStripMenuItem_Click(object sender, EventArgs e) // cut tool menu event
        { // open event 
            try //  try to cut text
            { //  open try block
                textBoxEditor.Cut(); // use cut method
            } // close try block
            catch // can't cut text
            { // open catch block
                _ = MessageBox.Show("Unable to cut document contents.", "RTP - Cut", MessageBoxButtons.OK, MessageBoxIcon.Error); //  error message
            } // close catch block
        } //  close event
        private void PasteToolStripMenuItem_Click(object sender, EventArgs e) // paste menu event
        { // open event
            try // try to paste text
            { // open try block
                textBoxEditor.Paste(); // paste method
            } // close try block
            catch // can't paste
            { // open catch block
                _ = MessageBox.Show("Unable to copy clipboard contents to document.", "RTP - Paste", MessageBoxButtons.OK, MessageBoxIcon.Error); // error message
            } // close catch block
        } // close event
        private void UndoToolStripMenuItem_Click(object sender, EventArgs e) // undo menu event
        { //  open event
            try //  try to undo the last command
            { // open try block
                if (textBoxEditor.CanUndo) textBoxEditor.Undo(); // undo method
            } // close try block
            catch // can't undo
            { // open catch block
                SystemSounds.Hand.Play(); // error sound
            } // close catch block
        } // close event
        private void RedoToolStripMenuItem_Click(object sender, EventArgs e) // redo menu event
        { //  open event
            try // try to redo
            { //  open try block
                if (textBoxEditor.CanRedo) textBoxEditor.Redo(); // if you can use the redo method in the editor, use it
            } // close try block
            catch // can't use redo
            { // open catch block
                SystemSounds.Hand.Play(); // player error sound
            } // close catch block
        } // close event
        #endregion // end of edit menu


        #region Publish Menu Methods

        private void PreviewToolStripMenuItem_Click(object sender, EventArgs e) // preview menu event
        { // open event
            try // try to dispay the print preview dialog
            { // open try block
                printPreviewDialog.Document = printDocument; // set the printing document
                _ = printPreviewDialog.ShowDialog();         // invoke print preview dialog method
            } // close try block
            catch // can't use print preview
            { // open catch block
                SystemSounds.Hand.Play(); // error sound
            } // close catch block
        } // close event
        private void PrintToolStripMenuItem_Click(object sender, EventArgs e) // print menu event
        { // open event 
            try // try to print
            { // open try block
                printDialog.Document = printDocument;                                   // set printing document
                if (printDialog.ShowDialog() == DialogResult.OK) printDocument.Print(); // if the user clicks ok on the print dialog, use the print method
            } // close try block
            catch // can't print
            { // open catch block
                SystemSounds.Hand.Play(); // error sound
            } // close catch block
        } // close event
        private void PageSetupToolStripMenuItem_Click(object sender, EventArgs e) // page setup menu event
        { // open event
            try // try to open print setup
            { //  open try block
                pageSetupDialog.Document = printDocument; // set printing document
                _ = pageSetupDialog.ShowDialog(); // use print dialog method
            } // close try block
            catch // can't use print setup
            { // open catch block
                SystemSounds.Hand.Play(); // error sound
            } // close catch block
        } // close event
        private void PrintDocument_BeginPrint(object sender, PrintEventArgs e) // print document method 1
        { // open method
            char[] param = { '\n' }; // get an array of the numder of new lines
            lines = printDialog.PrinterSettings.PrintRange == PrintRange.Selection // set up printing based on document parameters
                ? textBoxEditor.SelectedText.Split(param)
                : textBoxEditor.Text.Split(param);
            int i = 0;
            char[] trimParam = { '\r' };
            foreach (string s in lines) lines[i++] = s.TrimEnd(trimParam);
        } // close method
        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e) // print document method 2
        { // open method
            int x = e.MarginBounds.Left;                           // set margins
            int y = e.MarginBounds.Top;
            Brush brush = new SolidBrush(textBoxEditor.ForeColor); // print with color
            while (linesPrinted < lines.Length) // loop for whole document
            { // open while loop
                e.Graphics.DrawString(lines[linesPrinted++], // draw the line
                    textBoxEditor.Font, brush, x, y);        // print with font
                y += 15;                                     // increase y
                if (y >= e.MarginBounds.Bottom) // if y is higher than the bottom of the page
                { // open if statement
                    e.HasMorePages = true; // document has more pages
                    return;
                } // close if statement
            } // close while loop
            linesPrinted = 0;       // reset number of lines printed
            e.HasMorePages = false; // end of document
        } // close method
        #endregion // end of publish menu


        #region Format Tool Methods

        private void ButtonFontSelect_Click(object sender, EventArgs e) // select font button event
        { // open event
            try // try to set current font
            { // open try block
                fontDialog.Font = textBoxEditor.SelectionFont ?? null;                                         // logic for font diaglog box
                fontDialog.ShowApply = true;
                if (fontDialog.ShowDialog() == DialogResult.OK) textBoxEditor.SelectionFont = fontDialog.Font; // if user clicks ok, set selected font
            } //  close try block
            catch // can't select font
            { //  open catch block
                SystemSounds.Hand.Play(); // play error sound
            } // close catch block
        } // close event
        private void ButtonFontColor_Click(object sender, EventArgs e) // font color button event
        { // open event 
            try // try to select font color
            { // open try block
                colorDialog.Color = textBoxEditor.ForeColor;                                                       // color dialog
                if (colorDialog.ShowDialog() == DialogResult.OK) textBoxEditor.SelectionColor = colorDialog.Color; // user clicks ok, set selected color
            }// close try block
            catch // cant' select font color
            { // open catch block
                SystemSounds.Hand.Play(); // error sound
            } // close catch block
        } // close event
        private void ButtonBold_Click(object sender, EventArgs e) // bold button event
        { // open event
            try // try to make text bold
            { // open try block
                if (textBoxEditor.SelectionFont != null) // if font is not set to null
                { // open if
                    Font currentFont = textBoxEditor.SelectionFont;                                                 // set current font
                    FontStyle newFontStyle;                                                                         // change font style
                    newFontStyle = textBoxEditor.SelectionFont.Style ^ FontStyle.Bold;                              // set to bold
                    textBoxEditor.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle); // pass to editor
                } // close if
            } // close try block
            catch // can't set to bold
            { // open catch block
                SystemSounds.Hand.Play(); //  play error sound
            } // close catch block
        } // close event
        private void ButtonItalic_Click(object sender, EventArgs e) // italics button event
        { // open event
            try // try to make text italics
            { // open try block
                if (textBoxEditor.SelectionFont != null) // if font is not set to null 
                { // open if 
                    Font currentFont = textBoxEditor.SelectionFont;                                                 // set current font
                    FontStyle newFontStyle;                                                                         // change font style
                    newFontStyle = textBoxEditor.SelectionFont.Style ^ FontStyle.Italic;                            // set to italics
                    textBoxEditor.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle); // pass to editor
                } // close if
            } // close try block
            catch // can't italicize
            { // open catch block
                SystemSounds.Hand.Play(); //  play error sound
            } // close catch block
        } // close event
        private void ButtonUnderline_Click(object sender, EventArgs e) // underline button event
        { // open event
            try // try to underline text
            { // open try block
                if (textBoxEditor.SelectionFont != null) // if font is not set to null 
                { // open if
                    Font currentFont = textBoxEditor.SelectionFont;                                                 // set current font
                    FontStyle newFontStyle;                                                                         // change font style
                    newFontStyle = textBoxEditor.SelectionFont.Style ^ FontStyle.Underline;                         // set to underline
                    textBoxEditor.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle); // pass to editor
                } // close if
            } // close try block
            catch // can't underline
            { // open catch block
                SystemSounds.Hand.Play(); // error sound
            } // close catch block
        } // close event
        private void ButtonAlignLeft_Click(object sender, EventArgs e) // align left button event
        { // open event
            try // try to align text left
            { // open try block
                textBoxEditor.SelectionAlignment = HorizontalAlignment.Left; // set alignment to left
                buttonAlignCenter.Checked = false;                           // set other alignment buttons to off
                buttonAlignRight.Checked = false;
            } // close try block
            catch // can't align text
            { // open catch block
                SystemSounds.Hand.Play(); // error sound
            } // close catch block
        } // close event
        private void ButtonAlignCenter_Click(object sender, EventArgs e) // align center button event
        { // open event
            try // try to align text to center
            { // open try block
                textBoxEditor.SelectionAlignment = HorizontalAlignment.Center; // set alignment to center
                buttonAlignRight.Checked = false;                              // set other alignment buttons to off
                buttonAlignLeft.Checked = false;
            } // close try block
            catch // can't align text
            { // open catch block
                SystemSounds.Hand.Play(); // error sound
            } // close catch block
        } // close event
        private void ButtonAlignRight_Click(object sender, EventArgs e) // align right button event
        { // open event
            try // try to align text to right
            { // open try block
                textBoxEditor.SelectionAlignment = HorizontalAlignment.Right; // set alignment to right
                buttonAlignCenter.Checked = false;                            // set other alignment buttons to off
                buttonAlignLeft.Checked = false;
            } // close try block
            catch // can't align text
            { // open catch block
                SystemSounds.Hand.Play(); // error sound
            } // close catch block
        } // close event
        private void ButtonBullets_Click(object sender, EventArgs e) // bulleted text button event
        { // open event
            if (textBoxEditor.SelectionBullet == false) // if the text is not bulleted already
            { // open if
                try // try to add bullets
                { // open try block
                    textBoxEditor.BulletIndent = 10;      // set indention
                    textBoxEditor.SelectionBullet = true; // add bullets
                } // close try block
                catch // can't bullet text
                { // open catch block
                    SystemSounds.Hand.Play(); // error sound
                } // close catch block
            } // close if
            else // otherwise text is already bulleted
            { // open else
                try // try to remove bullets
                { // open try block
                    textBoxEditor.SelectionBullet = false; // set to unbulleted
                } // close try block
                catch // can't remove bullets
                { // open catch block
                    SystemSounds.Hand.Play(); // error sound
                } // close catch block
            } // close else
        } // close event
        #endregion // end format toolbar


        #region Form Closing Handler

        private void Main_FormClosing(object sender, FormClosingEventArgs e) // form closing handler method
        { // open method
            try // try to close
            { // open try block
                if (textBoxEditor.Modified) // if the editor has been modified
                { // open if
                    DialogResult answer = MessageBox.Show(       // open message box
                        "Save current document before exiting?", // asking user if they want to save
                        "Unsaved Document",                      // header
                        MessageBoxButtons.YesNo,                 // yes and no buttons
                        MessageBoxIcon.Question);
                    if (answer == DialogResult.No) // if user answers no
                    { // open nest if
                        textBoxEditor.Modified = false; // set editor unmodified
                        textBoxEditor.Clear();          // clear editor
                        return;
                    } // close nested if
                    else SaveToolStripMenuItem_Click(this, new EventArgs()); // save menu event method
                } // close if
                else textBoxEditor.Clear();                 // clear editor
                currentDoc = "";                            // empty current document string
            } // close try block
            catch // can't close form
            { // open catch block
                SystemSounds.Hand.Play(); // play error sound
            } // close catch block
        } // close method
        #endregion


        #region Word and Character Count

        private void TextChanged(object sender, EventArgs e)
        {
            TextChanged_WordCount(sender, e);
            TextChanged_CharacterCount(sender, e);
        }
        private void TextChanged_WordCount(object sender, EventArgs e)
        {
            string[] words = Regex.Split(textBoxEditor.Text, @"\W+");
            int wordCount = words.Length;
            labelWordCount.Text = wordCount == 0 || wordCount > 1 ? $"{wordCount} words" : "1 word";
        }
        private void TextChanged_CharacterCount(object sender, EventArgs e)
        {
            int charCount = textBoxEditor.Text.Length;
            labelCharCount.Text = charCount == 0 || charCount > 1 ? $"{charCount} characters" : "1 character";
        }
        #endregion

    } // close form
} // close namespace