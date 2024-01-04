using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Media;
using System.Windows.Forms;

namespace Rich_Text_Processor
{
    public static class PublishMenuHandler
    {
        public static string[] Lines { get; set; }
        private static int LinesPrinted { get; set; }

        public static void HandlePreview(PrintPreviewDialog printPreviewDialog, PrintDocument printDocument)
        {
            try
            {
                printPreviewDialog.Document = printDocument;
                printPreviewDialog.ShowDialog();
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling Print Preview: {ex.Message}");
                Messager.ShowErrorMessage("Error occurred during Print Preview.", "Print Preview Error");
                SystemSounds.Hand.Play();
            }
        }

        public static void HandlePrint(PrintDialog printDialog, PrintDocument printDocument)
        {
            try
            {
                printDialog.Document = printDocument;
                if (printDialog.ShowDialog() == DialogResult.OK) printDocument.Print();
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling Print: {ex.Message}");
                Messager.ShowErrorMessage("Error occurred during Print.", "Print Error");
                SystemSounds.Hand.Play();
            }
        }

        public static void HandlePageSetup(PageSetupDialog pageSetupDialog, PrintDocument printDocument)
        {
            try
            {
                pageSetupDialog.Document = printDocument;
                pageSetupDialog.ShowDialog();
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling Page Setup: {ex.Message}");
                Messager.ShowErrorMessage("Error occurred during Page Setup.", "Page Setup Error");
                SystemSounds.Hand.Play();
            }
        }

        public static void HandleBeginPrint(MagicSpellBox magicSpellBox, PrintDialog printDialog)
        {
            try
            {
                if (printDialog.PrinterSettings.PrintRange == PrintRange.Selection) Lines = magicSpellBox.SelectedText.Split(new char[] { '\n' });
                else Lines = magicSpellBox.Text.Split(new char[] { '\n' });
                int i = 0;
                foreach (string s in Lines) Lines[i++] = s.TrimEnd(new char[] { '\r' });
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling Begin Print: {ex.Message}");
                Messager.ShowErrorMessage("Error occurred during Begin Print.", "Begin Print Error");
                SystemSounds.Hand.Play();
            }
        }

        public static void HandlePrintPage(MagicSpellBox magicSpellBox, PrintPageEventArgs e)
        {
            try
            {
                int x = e.MarginBounds.Left;
                int y = e.MarginBounds.Top;
                using (Brush brush = new SolidBrush(magicSpellBox.ForeColor))
                {
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
                }

                LinesPrinted = 0;
                e.HasMorePages = false;
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling Print Page: {ex.Message}");
                Messager.ShowErrorMessage("Error occurred during Print Page.", "Print Page Error");
                SystemSounds.Hand.Play();
            }
        }
    }
}