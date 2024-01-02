using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Media;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Forms;

namespace Rich_Text_Processor
{ // open namespace
    public partial class MainForm : Form
    { // open class
        public MainForm() => InitializeComponent();

        private string CurrentFile { get; set; }
        private int LinesPrinted { get; set; }
        private string[] Lines { get; set; }

        #region File Menu Methods

        private void NewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (magicSpellBox.Modified)
                {
                    DialogResult answer;
                    answer = System.Windows.Forms.MessageBox.Show("Save current document before creating new document?",
                                                                "Unsaved Document",
                                                                MessageBoxButtons.YesNo,
                                                                MessageBoxIcon.Question);
                    if (answer == DialogResult.No)
                    {
                        CurrentFile = "";
                        Text = "Editor: New Document";
                        magicSpellBox.Modified = false;
                        magicSpellBox.ResetText();
                        return;
                    }
                    else
                    {
                        SaveToolStripMenuItem_Click(this, new EventArgs());
                        magicSpellBox.Modified = false;
                        magicSpellBox.ResetText();
                        CurrentFile = "";
                        Text = "Editor: New Document";
                        return;
                    }
                }
                else
                {
                    CurrentFile = "";
                    Text = "Editor: New Document";
                    magicSpellBox.Modified = false;
                    magicSpellBox.ResetText();
                    return;
                }
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (magicSpellBox.Modified)
                {
                    DialogResult answer;
                    answer = System.Windows.Forms.MessageBox.Show("Save current file before opening another document?",
                                                                "Unsaved Document",
                                                                MessageBoxButtons.YesNo,
                                                                MessageBoxIcon.Question);
                    if (answer == DialogResult.No)
                    {
                        magicSpellBox.Modified = false;
                        OpenFile();
                    }
                    else
                    {
                        SaveToolStripMenuItem_Click(this, new EventArgs());
                        OpenFile();
                    }
                }
                else OpenFile();
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void OpenFile()
        {
            try
            {
                openFileDialog.Title = "RTP - Open File";
                openFileDialog.DefaultExt = "rtf";
                openFileDialog.Filter = "Rich Text Files|*.rtf|Text Files|*.txt|HTML Files|*.htm|All Files|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.FileName = string.Empty;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (openFileDialog.FileName == "") return;

                    string strExt = Path.GetExtension(saveFileDialog.FileName).ToUpper();
                    if (strExt == ".RTF")
                    {
                        var textRange = new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd); // Load RTF content into the MagicSpellBox
                    }
                    else
                    {
                        using (StreamReader txtReader = new StreamReader(openFileDialog.FileName)) magicSpellBox.Text = txtReader.ReadToEnd();
                        TextRange textRange = new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd); // Set the selection range to the entire text
                        magicSpellBox.Box.Selection.Select(textRange.Start, textRange.End);
                    }
                    CurrentFile = openFileDialog.FileName;
                    magicSpellBox.Modified = false;
                    Text = $"Editor: {CurrentFile}";
                }
                else System.Windows.Forms.MessageBox.Show("Open File request cancelled by user.", "Cancelled");
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                saveFileDialog.Title = "RTP - Save File";
                saveFileDialog.DefaultExt = "rtf";
                saveFileDialog.Filter = "Rich Text Files|*.rtf|Text Files|*.txt|HTML Files|*.htm|All Files|*.*";
                saveFileDialog.FilterIndex = 1;
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (saveFileDialog.FileName == "") return;

                    string strExt = Path.GetExtension(saveFileDialog.FileName).ToUpper();
                    if (strExt == ".RTF")
                    {
                        var textRange = new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd);                      // Save RTF content from MagicSpellBox
                        using (var fs = new FileStream(saveFileDialog.FileName, FileMode.Create)) textRange.Save(fs, System.Windows.Forms.DataFormats.Rtf);
                    }
                    else
                    {
                        using (StreamWriter txtWriter = new StreamWriter(saveFileDialog.FileName)) txtWriter.Write(magicSpellBox.Text);                     // Save plain text content from MagicSpellBox
                        magicSpellBox.Box.Selection.Select(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentStart);
                    }
                    CurrentFile = saveFileDialog.FileName;
                    magicSpellBox.Modified = false;
                    Text = $"Editor: {CurrentFile}";
                    System.Windows.Forms.MessageBox.Show(CurrentFile.ToString() + " saved.", "File Save");
                }
                else System.Windows.Forms.MessageBox.Show("Save File request cancelled by user.", "Cancelled");
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void SaveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                saveFileDialog.Title = "RTP - Save File";
                saveFileDialog.DefaultExt = "rtf";
                saveFileDialog.Filter = "Rich Text Files|*.rtf|Text Files|*.txt|HTML Files|*.htm|All Files|*.*";
                saveFileDialog.FilterIndex = 1;
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (saveFileDialog.FileName == "") return;

                    string strExt = Path.GetExtension(saveFileDialog.FileName).ToUpper();
                    if (strExt == ".RTF")
                    {
                        var textRange = new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd);                                                                          // Save RTF content from the MagicSpellBox
                        using (var fs = new FileStream(saveFileDialog.FileName, FileMode.Create)) textRange.Save(fs, System.Windows.Forms.DataFormats.Rtf);
                    }
                    else using (var txtWriter = new StreamWriter(saveFileDialog.FileName)) txtWriter.Write(new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd).Text); // Save plain text content from the MagicSpellBox
                    magicSpellBox.Box.Selection.Select(magicSpellBox.Box.Document.ContentEnd, magicSpellBox.Box.Document.ContentEnd);                                                                           // Clear the selection
                    CurrentFile = saveFileDialog.FileName;
                    magicSpellBox.Modified = false;
                    Text = $"Editor: {CurrentFile}";
                    System.Windows.Forms.MessageBox.Show(CurrentFile.ToString() + " saved.", "File Save");
                }
                else System.Windows.Forms.MessageBox.Show("Save File request cancelled by user.", "Cancelled");
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (magicSpellBox.Modified == true)
                {
                    DialogResult answer;
                    answer = System.Windows.Forms.MessageBox.Show("Save this document before closing?",
                                                                "Unsaved Document",
                                                                MessageBoxButtons.YesNo,
                                                                MessageBoxIcon.Question);
                    if (answer == DialogResult.Yes) return;
                    else
                    {
                        magicSpellBox.Modified = false;
                        System.Windows.Forms.Application.Exit();
                    }
                }
                else
                {
                    magicSpellBox.Modified = false;
                    System.Windows.Forms.Application.Exit();
                }
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        #endregion

        #region Edit Menu Methods

        private void SelectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                magicSpellBox.SelectAll();
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Unable to select all document content.",
                                                    "RTE - Select",
                                                    MessageBoxButtons.OK,
                                                    MessageBoxIcon.Error);
            }
        }

        private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                magicSpellBox.Copy();
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Unable to copy document content.",
                                                    "RTE - Copy",
                                                    MessageBoxButtons.OK,
                                                    MessageBoxIcon.Error);
            }
        }

        private void CutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                magicSpellBox.Cut();
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Unable to cut document content.",
                                                    "RTE - Cut",
                                                    MessageBoxButtons.OK,
                                                    MessageBoxIcon.Error);
            }
        }

        private void PasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                magicSpellBox.Paste();
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Unable to copy clipboard content to document.",
                                                    "RTE - Paste",
                                                    MessageBoxButtons.OK,
                                                    MessageBoxIcon.Error);
            }
        }

        private void UndoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (magicSpellBox.CanUndo) magicSpellBox.Undo();
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void RedoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (magicSpellBox.CanRedo) magicSpellBox.Redo();
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        #endregion // end of edit menu

        #region Publish Menu Methods

        private void PreviewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                printPreviewDialog.Document = printDocument;
                printPreviewDialog.ShowDialog();
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void PrintToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                printDialog.Document = printDocument;
                if (printDialog.ShowDialog() == DialogResult.OK) printDocument.Print();
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void PageSetupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                pageSetupDialog.Document = printDocument;
                pageSetupDialog.ShowDialog();
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void PrintDocument_BeginPrint(object sender, PrintEventArgs e)
        {
            char[] param = { '\n' };

            if (printDialog.PrinterSettings.PrintRange == PrintRange.Selection) Lines = magicSpellBox.SelectedText.Split(param);
            else Lines = magicSpellBox.Text.Split(param);
            int i = 0;
            char[] trimParam = { '\r' };
            foreach (string s in Lines) Lines[i++] = s.TrimEnd(trimParam);
        }

        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            int x = e.MarginBounds.Left;
            int y = e.MarginBounds.Top;
            Brush brush = new SolidBrush(magicSpellBox.ForeColor);

            while (LinesPrinted < Lines.Length)
            {
                e.Graphics.DrawString(Lines[LinesPrinted++], magicSpellBox.Font, brush, x, y);
                y += 15;
                if (y >= e.MarginBounds.Bottom)
                {
                    e.HasMorePages = true;
                    return;
                }
            }
            LinesPrinted = 0;
            e.HasMorePages = false;
        }

        #endregion // end of publish menu

        #region Format Tool Methods

        private void ButtonFontSelect_Click(object sender, EventArgs e)
        {
            try
            {
                if (magicSpellBox.SelectionFont != null) fontDialog.Font = magicSpellBox.SelectionFont;
                else fontDialog.Font = null;
                fontDialog.ShowApply = true;
                if (fontDialog.ShowDialog() == DialogResult.OK) magicSpellBox.SelectionFont = fontDialog.Font;
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void ButtonFontColor_Click(object sender, EventArgs e)
        {
            try
            {
                colorDialog.Color = magicSpellBox.ForeColor;
                if (colorDialog.ShowDialog() == DialogResult.OK) magicSpellBox.ApplySelectionForeground(colorDialog.Color); // Apply the color to the current selection if possible
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void ButtonBold_Click(object sender, EventArgs e)
        {
            try
            {
                magicSpellBox.Bold();
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void ButtonItalic_Click(object sender, EventArgs e)
        {
            try
            {
                magicSpellBox.Italic();
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void ButtonUnderline_Click(object sender, EventArgs e)
        {
            try
            {
                magicSpellBox.Underline();
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void ButtonAlignLeft_Click(object sender, EventArgs e)
        {
            try
            {
                magicSpellBox.SetAlignment(TextAlignment.Left);
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void ButtonAlignCenter_Click(object sender, EventArgs e)
        {
            try
            {
                magicSpellBox.SetAlignment(TextAlignment.Center);
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void ButtonAlignRight_Click(object sender, EventArgs e)
        {
            try
            {
                magicSpellBox.SetAlignment(TextAlignment.Right);
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void ButtonBullets_Click(object sender, EventArgs e)
        {
            try
            {
                magicSpellBox.Bullet();
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        #endregion // end format toolbar

        #region Form Closing Handler

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (magicSpellBox.Modified == true)
                {
                    DialogResult answer;
                    answer = System.Windows.Forms.MessageBox.Show("Save current document before exiting?",
                                                                "Unsaved Document",
                                                                MessageBoxButtons.YesNo,
                                                                MessageBoxIcon.Question);
                    if (answer == DialogResult.No)
                    {
                        magicSpellBox.Modified = false;
                        magicSpellBox.ResetText();
                        return;
                    }
                    else SaveToolStripMenuItem_Click(this, new EventArgs());
                }
                else magicSpellBox.ResetText();
                CurrentFile = "";
                Text = "Editor: New Document";
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        #endregion // end form closing handler

        #region Word and Character Count

        private new void TextChanged(object sender, EventArgs e)
        {
            TextChanged_WordCount();
            TextChanged_CharacterCount();
        }

        private void TextChanged_WordCount()
        {
            string text = magicSpellBox.Text.Trim();

            if (string.IsNullOrEmpty(text)) // Check if the text is empty
            {
                labelWordCount.Text = "0 words";
                return;
            }

            string[] words = text.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            labelWordCount.Text = words.Length > 1 ? $"{words.Length} words" : "1 word";
        }

        private void TextChanged_CharacterCount()
        {
            TextRange textRange = new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd);
            labelCharCount.Text = $"{textRange.Text.Count(c => !char.IsWhiteSpace(c))} characters";
        }

        #endregion // end word and character count
    } // close class
} // close namespace