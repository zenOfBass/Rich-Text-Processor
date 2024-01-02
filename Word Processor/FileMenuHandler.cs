using System.IO;
using System.Media;
using System.Windows.Documents;
using System.Windows.Forms;

namespace Rich_Text_Processor
{
    public static class FileMenuHandler
    {
        public static void HandleNew(MainForm form, MagicSpellBox magicSpellBox, SaveFileDialog saveFileDialog)
        {
            try
            {
                if (magicSpellBox.Modified && string.IsNullOrEmpty(magicSpellBox.Text.Trim()))
                {
                    DialogResult answer = MessageBox.Show("Save current document before creating new document?",
                                                        "Unsaved Document",
                                                        MessageBoxButtons.YesNo,
                                                        MessageBoxIcon.Question);

                    if (answer == DialogResult.Yes) HandleSave(form, magicSpellBox, saveFileDialog);
                }

                form.CurrentFile = "";
                form.Text = "Editor: New Document";
                magicSpellBox.Modified = false;
                magicSpellBox.ResetText();
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        public static void HandleOpen(MainForm form, MagicSpellBox magicSpellBox, SaveFileDialog saveFileDialog, OpenFileDialog openFileDialog)
        {
            try
            {
                if (magicSpellBox.Modified && string.IsNullOrEmpty(magicSpellBox.Text.Trim()))
                {
                    DialogResult answer = MessageBox.Show("Save current file before opening another document?",
                                                        "Unsaved Document",
                                                        MessageBoxButtons.YesNo,
                                                        MessageBoxIcon.Question);

                    if (answer == DialogResult.Yes) HandleSave(form, magicSpellBox, saveFileDialog);
                }

                OpenFile(form, magicSpellBox, openFileDialog);
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        public static void HandleSave(MainForm form, MagicSpellBox magicSpellBox, SaveFileDialog saveFileDialog)
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
                        var textRange = new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd);
                        using (var fs = new FileStream(saveFileDialog.FileName, FileMode.Create)) textRange.Save(fs, DataFormats.Rtf);
                    }
                    else
                    {
                        using (StreamWriter txtWriter = new StreamWriter(saveFileDialog.FileName)) txtWriter.Write(magicSpellBox.Text);
                        magicSpellBox.Box.Selection.Select(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentStart);
                    }

                    form.CurrentFile = saveFileDialog.FileName;
                    magicSpellBox.Modified = false;
                    form.Text = $"Editor: {form.CurrentFile}";
                    MessageBox.Show($"{form.CurrentFile} saved.", "File Save");
                }
                else MessageBox.Show("Save File request canceled by user.", "Canceled");
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        public static void HandleExit(MainForm form, MagicSpellBox magicSpellBox, SaveFileDialog saveFileDialog)
        {
            try
            {
                if (magicSpellBox.Modified && string.IsNullOrEmpty(magicSpellBox.Text.Trim()))
                {
                    DialogResult answer = MessageBox.Show("Save this document before closing?",
                                                        "Unsaved Document",
                                                        MessageBoxButtons.YesNo,
                                                        MessageBoxIcon.Question);

                    if (answer == DialogResult.Yes) HandleSave(form, magicSpellBox, saveFileDialog);
                    else
                    {
                        magicSpellBox.Modified = false;
                        Application.Exit();
                    }
                }
                else
                {
                    magicSpellBox.Modified = false;
                    Application.Exit();
                }
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        private static void OpenFile(MainForm form, MagicSpellBox magicSpellBox, OpenFileDialog openFileDialog)
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

                    string strExt = Path.GetExtension(openFileDialog.FileName).ToUpper();

                    if (strExt == ".RTF")
                    {
                        var textRange = new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd);
                    }
                    else
                    {
                        using (StreamReader txtReader = new StreamReader(openFileDialog.FileName)) magicSpellBox.Text = txtReader.ReadToEnd();
                        TextRange textRange = new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd);
                        magicSpellBox.Box.Selection.Select(textRange.Start, textRange.End);
                    }

                    form.CurrentFile = openFileDialog.FileName;
                    magicSpellBox.Modified = false;
                    form.Text = $"Editor: {form.CurrentFile}";
                }
                else MessageBox.Show("Open File request canceled by user.", "Canceled");
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }
    }
}