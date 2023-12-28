using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Media;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Forms;

namespace Rich_Text_Processor // this program will be a word processor based around the rich text box
{ // open namespace
    public partial class MainWindow : Form
    { // open class
        public MainWindow() => InitializeComponent();

        private string currentFile;
        private int linesPrinted;
        private string[] lines;

        #region File Menu Methods

        private void NewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (magicSpellBox.Modified == true)
                {
                    DialogResult answer;
                    answer = System.Windows.Forms.MessageBox.Show("Save current document before creating new document?", "Unsaved Document", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (answer == DialogResult.No)
                    {
                        currentFile = "";
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
                        currentFile = "";
                        Text = "Editor: New Document";
                        return;
                    }
                }
                else
                {
                    currentFile = "";
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
                if (magicSpellBox.Modified == true)
                {
                    DialogResult answer;
                    answer = System.Windows.Forms.MessageBox.Show("Save current file before opening another document?", "Unsaved Document", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
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
                else
                {
                    OpenFile();
                }
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
                    string strExt;
                    strExt = Path.GetExtension(openFileDialog.FileName);
                    strExt = strExt.ToUpper();
                    if (strExt == ".RTF")
                    {
                        // Load RTF content into the MagicSpellBox
                        var textRange = new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd);
                        using (var fs = new FileStream(openFileDialog.FileName, FileMode.Open))
                        {
                            textRange.Load(fs, System.Windows.Forms.DataFormats.Rtf);
                        }
                    }
                    else
                    {
                        StreamReader txtReader;
                        txtReader = new StreamReader(openFileDialog.FileName);
                        magicSpellBox.Text = txtReader.ReadToEnd();
                        txtReader.Close();
                        txtReader = null;

                        // Set the selection range to the entire text
                        TextRange textRange = new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd);
                        magicSpellBox.Box.Selection.Select(textRange.Start, textRange.End);
                    }
                    currentFile = openFileDialog.FileName;
                    magicSpellBox.Modified = false;
                    Text = "Editor: " + currentFile.ToString();
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
                    string strExt;
                    strExt = Path.GetExtension(saveFileDialog.FileName);
                    strExt = strExt.ToUpper();
                    if (strExt == ".RTF")
                    {
                        // Save RTF content from MagicSpellBox
                        var textRange = new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd);
                        using (var fs = new FileStream(saveFileDialog.FileName, FileMode.Create))
                        {
                            textRange.Save(fs, System.Windows.Forms.DataFormats.Rtf);
                        }
                    }
                    else
                    {
                        // Save plain text content from MagicSpellBox
                        StreamWriter txtWriter;
                        txtWriter = new StreamWriter(saveFileDialog.FileName);
                        txtWriter.Write(magicSpellBox.Text);
                        txtWriter.Close();
                        txtWriter = null;
                        magicSpellBox.Box.Selection.Select(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentStart);
                    }
                    currentFile = saveFileDialog.FileName;
                    magicSpellBox.Modified = false;
                    Text = "Editor: " + currentFile.ToString();
                    System.Windows.Forms.MessageBox.Show(currentFile.ToString() + " saved.", "File Save");
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
                    string strExt;
                    strExt = Path.GetExtension(saveFileDialog.FileName);
                    strExt = strExt.ToUpper();
                    if (strExt == ".RTF")
                    {
                        // Save RTF content from the MagicSpellBox
                        var textRange = new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd);
                        using (var fs = new FileStream(saveFileDialog.FileName, FileMode.Create))
                        {
                            textRange.Save(fs, System.Windows.Forms.DataFormats.Rtf);
                        }
                    }
                    else
                    {
                        // Save plain text content from the MagicSpellBox
                        using (var txtWriter = new StreamWriter(saveFileDialog.FileName))
                        {
                            txtWriter.Write(new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd).Text);
                        }
                    }

                    // Clear the selection
                    magicSpellBox.Box.Selection.Select(magicSpellBox.Box.Document.ContentEnd, magicSpellBox.Box.Document.ContentEnd);
                    currentFile = saveFileDialog.FileName;
                    magicSpellBox.Modified = false;
                    Text = "Editor: " + currentFile.ToString();
                    System.Windows.Forms.MessageBox.Show(currentFile.ToString() + " saved.", "File Save");
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
                    answer = System.Windows.Forms.MessageBox.Show("Save this document before closing?", "Unsaved Document", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
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
                System.Windows.Forms.MessageBox.Show("Unable to select all document content.", "RTE - Select", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                System.Windows.Forms.MessageBox.Show("Unable to copy document content.", "RTE - Copy", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                System.Windows.Forms.MessageBox.Show("Unable to cut document content.", "RTE - Cut", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                System.Windows.Forms.MessageBox.Show("Unable to copy clipboard content to document.", "RTE - Paste", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UndoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (magicSpellBox.CanUndo)
                {
                    magicSpellBox.Undo();
                }
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

            if (printDialog.PrinterSettings.PrintRange == PrintRange.Selection) lines = magicSpellBox.SelectedText.Split(param);
            else lines = magicSpellBox.Text.Split(param);
            int i = 0;
            char[] trimParam = { '\r' };
            foreach (string s in lines)
            {
                lines[i++] = s.TrimEnd(trimParam);
            }
        }

        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            int x = e.MarginBounds.Left;
            int y = e.MarginBounds.Top;
            System.Drawing.Brush brush = new SolidBrush(magicSpellBox.ForeColor);

            while (linesPrinted < lines.Length)
            {
                e.Graphics.DrawString(lines[linesPrinted++],
                    magicSpellBox.Font, brush, x, y);
                y += 15;
                if (y >= e.MarginBounds.Bottom)
                {
                    e.HasMorePages = true;
                    return;
                }
            }
            linesPrinted = 0;
            e.HasMorePages = false;
        }

        #endregion // end of publish menu

        #region Format Tool Methods

        private void ButtonFontSelect_Click(object sender, EventArgs e)
        {
            try
            {
                if (magicSpellBox.SelectionFont != null)
                {
                    fontDialog.Font = magicSpellBox.SelectionFont;
                }
                else
                {
                    fontDialog.Font = null;
                }
                fontDialog.ShowApply = true;
                if (fontDialog.ShowDialog() == DialogResult.OK)
                {
                    magicSpellBox.SelectionFont = fontDialog.Font;
                }
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

                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    // Apply the color to the current selection if possible
                    magicSpellBox.ApplySelectionForeground(colorDialog.Color);
                }
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
                var selection = magicSpellBox.Box.Selection;

                if (selection != null && !selection.IsEmpty)
                {
                    var currentFont = new System.Drawing.FontFamily(selection.GetPropertyValue(TextElement.FontFamilyProperty).ToString());
                    var currentSize = selection.GetPropertyValue(TextElement.FontSizeProperty);
                    var currentWeight = selection.GetPropertyValue(TextElement.FontWeightProperty);

                    FontWeight newWeight = (currentWeight != null && currentWeight.Equals(FontWeights.Bold)) ? FontWeights.Normal : FontWeights.Bold;

                    magicSpellBox.Box.Selection.ApplyPropertyValue(TextElement.FontWeightProperty, newWeight);
                }
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
                var selection = magicSpellBox.Box.Selection;

                if (selection != null && !selection.IsEmpty)
                {
                    var currentFont = new System.Drawing.FontFamily(selection.GetPropertyValue(TextElement.FontFamilyProperty).ToString());
                    var currentSize = selection.GetPropertyValue(TextElement.FontSizeProperty);
                    var currentStyle = selection.GetPropertyValue(TextElement.FontStyleProperty);

                    System.Windows.FontStyle newStyle = (currentStyle != null && currentStyle.Equals(FontStyles.Italic)) ? FontStyles.Normal : FontStyles.Italic;

                    magicSpellBox.Box.Selection.ApplyPropertyValue(TextElement.FontStyleProperty, newStyle);
                }
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
                if (magicSpellBox.SelectionFont != null)
                {
                    TextRange selectedText = new TextRange(magicSpellBox.Box.Selection.Start, magicSpellBox.Box.Selection.End);
                    selectedText.ApplyPropertyValue(Inline.TextDecorationsProperty, TextDecorations.Underline);
                }
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
                buttonAlignCenter.Checked = false;
                buttonAlignRight.Checked = false;
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
                buttonAlignRight.Checked = false;
                buttonAlignLeft.Checked = false;
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
                buttonAlignCenter.Checked = false;
                buttonAlignLeft.Checked = false;
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private void ButtonBullets_Click(object sender, EventArgs e)
        {
            // Get the current selection
            TextRange selection = new TextRange(magicSpellBox.Box.Selection.Start, magicSpellBox.Box.Selection.End);

            // Check if the selection is empty
            if (!selection.IsEmpty)
            {
                // Split the selected text into lines
                string[] lines = selection.Text.Split('\n');

                // Clear the existing selection
                selection.Text = "";

                // Insert a bullet point at the beginning of each line
                foreach (string line in lines)
                {
                    // Check if the line is not empty
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        magicSpellBox.Box.CaretPosition.InsertTextInRun("\u2022 " + line.Trim());

                        // Move the caret to the end of the inserted text
                        magicSpellBox.Box.CaretPosition = magicSpellBox.Box.CaretPosition.GetPositionAtOffset(4 + line.Trim().Length);
                    }

                    // Insert a newline after each line
                    magicSpellBox.Box.CaretPosition.InsertParagraphBreak();
                }
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
                    answer = System.Windows.Forms.MessageBox.Show("Save current document before exiting?", "Unsaved Document", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (answer == DialogResult.No)
                    {
                        magicSpellBox.Modified = false;
                        magicSpellBox.ResetText();
                        return;
                    }
                    else SaveToolStripMenuItem_Click(this, new EventArgs());
                }
                else magicSpellBox.ResetText();
                currentFile = "";
                Text = "Editor: New Document";
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        #endregion

        #region Word and Character Count

        private new void TextChanged(object sender, EventArgs e)
        {
            TextChanged_WordCount(sender, e);
            TextChanged_CharacterCount(sender, e);
        }

        private void TextChanged_WordCount(object sender, EventArgs e)
        {
            string text = magicSpellBox.Text.Trim();

            // Check if the text is empty
            if (string.IsNullOrEmpty(text))
            {
                labelWordCount.Text = "0 words";
                return;
            }

            string[] words = text.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int wordCount = words.Length;
            labelWordCount.Text = wordCount > 1 ? $"{wordCount} words" : "1 word";
        }

        private void TextChanged_CharacterCount(object sender, EventArgs e)
        {
            int charCount = magicSpellBox.Text.Count(c => !char.IsWhiteSpace(c));
            labelCharCount.Text = charCount > 0 ? $"{charCount} characters" : "0 character";
        }

        #endregion

    } // close class
} // close namespace