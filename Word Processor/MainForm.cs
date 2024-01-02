﻿using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Media;
using System.Windows;
using System.Windows.Forms;
using Word_Processor;

namespace Rich_Text_Processor
{ // open namespace
    public partial class MainForm : Form
    { // open class
        public MainForm() => InitializeComponent();

        public string CurrentFile { get; set; }

        #region File Menu Methods

        private void NewToolStripMenuItem_Click(object sender, EventArgs e) => FileMenuHandler.HandleNew(this, magicSpellBox, saveFileDialog);
        private void OpenToolStripMenuItem_Click(object sender, EventArgs e) => FileMenuHandler.HandleOpen(this, magicSpellBox, saveFileDialog, openFileDialog);
        private void SaveToolStripMenuItem_Click(object sender, EventArgs e) => FileMenuHandler.HandleSave(this, magicSpellBox, saveFileDialog);
        private void ExitToolStripMenuItem_Click(object sender, EventArgs e) => FileMenuHandler.HandleExit(this, magicSpellBox, saveFileDialog);

        #endregion

        #region Edit Menu Methods

        private void SelectAllToolStripMenuItem_Click(object sender, EventArgs e) => EditMenuHandler.HandleSelectAll(magicSpellBox);
        private void CopyToolStripMenuItem_Click(object sender, EventArgs e) => EditMenuHandler.HandleCopy(magicSpellBox);
        private void CutToolStripMenuItem_Click(object sender, EventArgs e) => EditMenuHandler.HandleCut(magicSpellBox);
        private void PasteToolStripMenuItem_Click(object sender, EventArgs e) => EditMenuHandler.HandlePaste(magicSpellBox);
        private void UndoToolStripMenuItem_Click(object sender, EventArgs e) => EditMenuHandler.HandleUndo(magicSpellBox);
        private void RedoToolStripMenuItem_Click(object sender, EventArgs e) => EditMenuHandler.HandleRedo(magicSpellBox);

        #endregion // end of edit menu

        #region Publish Menu Methods

        private void PreviewToolStripMenuItem_Click(object sender, EventArgs e) => PublishMenuHandler.HandlePreview(printPreviewDialog, printDocument);
        private void PrintToolStripMenuItem_Click(object sender, EventArgs e) => PublishMenuHandler.HandlePrint(printDialog, printDocument);
        private void PageSetupToolStripMenuItem_Click(object sender, EventArgs e) => PublishMenuHandler.HandlePageSetup(pageSetupDialog, printDocument);
        private void PrintDocument_BeginPrint(object sender, PrintEventArgs e) => PublishMenuHandler.HandleBeginPrint(magicSpellBox, printDialog, printDocument, e);
        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e) => PublishMenuHandler.HandlePrintPage(magicSpellBox, e);

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

        #region Main Form Closing

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (magicSpellBox.Modified == true && string.IsNullOrEmpty(magicSpellBox.Text.Trim()))
                {
                    DialogResult answer = System.Windows.Forms.MessageBox.Show("Save current document before exiting?",
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

        private void TextChanged_WordCount() => labelWordCount.Text = string.IsNullOrEmpty(magicSpellBox.Text.Trim())
                ? "0 words" : magicSpellBox.WordCount <= 1 ? "1 word" : $"{magicSpellBox.WordCount} words";

        private void TextChanged_CharacterCount() => labelCharCount.Text = $"{magicSpellBox.CharCount} characters";

        #endregion // end word and character count
    } // close class
} // close namespace