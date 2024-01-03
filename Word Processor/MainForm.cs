using System;
using System.Drawing.Printing;
using System.Media;
using System.Windows.Forms;

namespace Rich_Text_Processor
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            magicSpellBox.Text = string.Empty;
        }

        public string CurrentFile { get; set; }

        #region File Menu

        private void NewToolStripMenuItem_Click(object sender, EventArgs e) => FileMenuHandler.HandleNew(this, magicSpellBox, saveFileDialog);
        private void OpenToolStripMenuItem_Click(object sender, EventArgs e) => FileMenuHandler.HandleOpen(this, magicSpellBox, saveFileDialog, openFileDialog);
        private void SaveToolStripMenuItem_Click(object sender, EventArgs e) => FileMenuHandler.HandleSave(this, magicSpellBox, saveFileDialog);
        private void ExitToolStripMenuItem_Click(object sender, EventArgs e) => FileMenuHandler.HandleExit(this, magicSpellBox, saveFileDialog);

        #endregion

        #region Edit Menu

        private void SelectAllToolStripMenuItem_Click(object sender, EventArgs e) => EditMenuHandler.HandleSelectAll(magicSpellBox);
        private void CopyToolStripMenuItem_Click(object sender, EventArgs e) => EditMenuHandler.HandleCopy(magicSpellBox);
        private void CutToolStripMenuItem_Click(object sender, EventArgs e) => EditMenuHandler.HandleCut(magicSpellBox);
        private void PasteToolStripMenuItem_Click(object sender, EventArgs e) => EditMenuHandler.HandlePaste(magicSpellBox);
        private void UndoToolStripMenuItem_Click(object sender, EventArgs e) => EditMenuHandler.HandleUndo(magicSpellBox);
        private void RedoToolStripMenuItem_Click(object sender, EventArgs e) => EditMenuHandler.HandleRedo(magicSpellBox);

        #endregion // end of edit menu

        #region Publish Menu

        private void PreviewToolStripMenuItem_Click(object sender, EventArgs e) => PublishMenuHandler.HandlePreview(printPreviewDialog, printDocument);
        private void PrintToolStripMenuItem_Click(object sender, EventArgs e) => PublishMenuHandler.HandlePrint(printDialog, printDocument);
        private void PageSetupToolStripMenuItem_Click(object sender, EventArgs e) => PublishMenuHandler.HandlePageSetup(pageSetupDialog, printDocument);
        private void PrintDocument_BeginPrint(object sender, PrintEventArgs e) => PublishMenuHandler.HandleBeginPrint(magicSpellBox, printDialog);
        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e) => PublishMenuHandler.HandlePrintPage(magicSpellBox, e);

        #endregion // end of publish menu

        #region Format Toolbar

        private void ButtonFontSelect_Click(object sender, EventArgs e) => FormatToolbarHandler.HandleFontSelect(magicSpellBox, fontDialog);
        private void ButtonFontColor_Click(object sender, EventArgs e) => FormatToolbarHandler.HandleFontColor(magicSpellBox, colorDialog);
        private void ButtonBold_Click(object sender, EventArgs e) => FormatToolbarHandler.HandleBold(magicSpellBox);
        private void ButtonItalic_Click(object sender, EventArgs e) => FormatToolbarHandler.HandleItalic(magicSpellBox);
        private void ButtonUnderline_Click(object sender, EventArgs e) => FormatToolbarHandler.HandleUnderline(magicSpellBox);
        private void ButtonAlignLeft_Click(object sender, EventArgs e) => FormatToolbarHandler.HandleAlignLeft(magicSpellBox);
        private void ButtonAlignCenter_Click(object sender, EventArgs e) => FormatToolbarHandler.HandleAlignCenter(magicSpellBox);
        private void ButtonAlignRight_Click(object sender, EventArgs e) => FormatToolbarHandler.HandleAlignRight(magicSpellBox);
        private void ButtonBullets_Click(object sender, EventArgs e) => FormatToolbarHandler.HandleBullets(magicSpellBox);

        #endregion // end format toolbar

        #region Status Strip

        private new void TextChanged(object sender, EventArgs e)
        {
            WordAndCharCountHandler.HandleWordCount(magicSpellBox, labelWordCount);
            WordAndCharCountHandler.HandleWordCount(magicSpellBox, labelCharCount);
        }

        #endregion // end status strip

        #region Hotkeys

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if ((keyData & Keys.Control) == Keys.Control)
            {
                switch (keyData & ~Keys.Control)
                {
                    case Keys.B:
                        FormatToolbarHandler.HandleBold(magicSpellBox);
                        return true;
                    case Keys.I:
                        FormatToolbarHandler.HandleItalic(magicSpellBox);
                        return true;
                    case Keys.U:
                        FormatToolbarHandler.HandleUnderline(magicSpellBox);
                        return true;
                    case Keys.L:
                        FormatToolbarHandler.HandleAlignLeft(magicSpellBox);
                        return true;
                    case Keys.M:
                        FormatToolbarHandler.HandleAlignCenter(magicSpellBox);
                        return true;
                    case Keys.R:
                        FormatToolbarHandler.HandleAlignRight(magicSpellBox);
                        return true;
                    case Keys.OemPeriod:
                        FormatToolbarHandler.HandleBullets(magicSpellBox);
                        return true;
                    case Keys.S:
                        FileMenuHandler.HandleSave(this, magicSpellBox, saveFileDialog);
                        return true;
                    case Keys.O:
                        FileMenuHandler.HandleOpen(this, magicSpellBox, saveFileDialog, openFileDialog);
                        return true;
                    case Keys.N:
                        FileMenuHandler.HandleNew(this, magicSpellBox, saveFileDialog);
                        return true;
                    case Keys.P:
                        PublishMenuHandler.HandlePrint(printDialog, printDocument);
                        return true;
                }
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        #endregion

        #region Main Form Closing

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (magicSpellBox.Modified && !string.IsNullOrEmpty(magicSpellBox.Text.Trim()))
                {
                    DialogResult answer = MessageBox.Show("Save current document before exiting?",
                                            "Unsaved Document",
                                            MessageBoxButtons.YesNo,
                                            MessageBoxIcon.Question);
                    if (answer == DialogResult.No)
                    {
                        magicSpellBox.Modified = false;
                        magicSpellBox.ResetText();
                        return;
                    }
                    else FileMenuHandler.HandleSave(this, magicSpellBox, saveFileDialog);
                }
                else magicSpellBox.ResetText();
                CurrentFile = "";
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        #endregion // end form closing handler
    }
}