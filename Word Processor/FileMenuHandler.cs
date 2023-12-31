﻿using System;
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
                if (magicSpellBox.Modified && !string.IsNullOrEmpty(magicSpellBox.Text.Trim()))
                    if (Messager.ShowYesNoMessage("Save the current document before creating a new document?", "Unsaved Document") == DialogResult.Yes) HandleSave(form, magicSpellBox, saveFileDialog);

                form.CurrentFile = "";
                magicSpellBox.Modified = false;
                magicSpellBox.ResetText();
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling new document: {ex.Message}");
                Messager.ShowErrorMessage($"Error handling new document: {ex.Message}", "Error");
                SystemSounds.Hand.Play();
            }
        }

        public static void HandleOpen(MainForm form, MagicSpellBox magicSpellBox, SaveFileDialog saveFileDialog, OpenFileDialog openFileDialog)
        {
            try
            {
                if (magicSpellBox.Modified && !string.IsNullOrEmpty(magicSpellBox.Text.Trim()))
                    if (Messager.ShowYesNoMessage("Save the current file before opening another document?", "Unsaved Document") == DialogResult.Yes) HandleSave(form, magicSpellBox, saveFileDialog);

                OpenFile(form, magicSpellBox, openFileDialog);
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling opening document: {ex.Message}");
                Messager.ShowErrorMessage($"Error handling opening document: {ex.Message}", "Error");
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
                    if (openFileDialog.FileName != "")
                    {
                        string strExt = Path.GetExtension(openFileDialog.FileName).ToUpper();

                        if (strExt == ".RTF")
                        {
                            try
                            {
                                using (FileStream fileStream = new FileStream(openFileDialog.FileName, FileMode.Open))
                                {
                                    var range = new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd);
                                    range.Load(fileStream, DataFormats.Rtf);
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.Log(LogLevel.Error, $"Error opening RTF file: {ex.Message}");
                                Messager.ShowErrorMessage($"Error opening RTF file: {ex.Message}", "Open File Error");
                                SystemSounds.Hand.Play();
                            }
                        }
                        else
                        {
                            using (StreamReader txtReader = new StreamReader(openFileDialog.FileName))
                            {
                                try
                                {
                                    magicSpellBox.Text = txtReader.ReadToEnd();
                                    TextRange textRange = new TextRange(magicSpellBox.Box.Document.ContentStart, magicSpellBox.Box.Document.ContentEnd);
                                    magicSpellBox.Box.Selection.Select(textRange.Start, textRange.End);
                                }
                                catch (Exception ex)
                                {
                                    Logger.Log(LogLevel.Error, $"Error opening file: {ex.Message}");
                                    Messager.ShowErrorMessage($"Error opening file: {ex.Message}", "Open File Error");
                                    SystemSounds.Hand.Play();
                                }
                            }
                        }

                        form.CurrentFile = openFileDialog.FileName;
                        magicSpellBox.Modified = false;
                        form.Text = $"Rich Text Processor: {form.CurrentFile}";
                    }
                }
                else Messager.ShowErrorMessage("Open File request canceled by user.", "Canceled");
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling open file: {ex.Message}");
                Messager.ShowErrorMessage($"Error handling open file: {ex.Message}", "Open File Error");
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
                    if (saveFileDialog.FileName != "")
                    {
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
                        form.Text = $"Rich Text Processor: {form.CurrentFile}";
                        MessageBox.Show($"{form.CurrentFile} saved.", "File Save");
                    }
                }
                else Messager.ShowErrorMessage("Save File request canceled by the user.", "Canceled");
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error saving document: {ex.Message}");
                Messager.ShowErrorMessage($"Error saving document: {ex.Message}", "Save Error");
                SystemSounds.Hand.Play();
            }
        }

        public static void HandleExit(MainForm form, MagicSpellBox magicSpellBox, SaveFileDialog saveFileDialog)
        {
            try
            {
                if (magicSpellBox.Modified && !string.IsNullOrEmpty(magicSpellBox.Text.Trim()) && Messager.ShowYesNoMessage("Save this document before closing?", "Unsaved Document") == DialogResult.Yes)
                    HandleSave(form, magicSpellBox, saveFileDialog);
                else
                {
                    magicSpellBox.Modified = false;
                    Application.Exit();
                }
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling exit: {ex.Message}");
                Messager.ShowErrorMessage($"Error handling exit: {ex.Message}", "Exit Error");
                SystemSounds.Hand.Play();
            }
        }
    }
}