using System.Drawing;
using System.Drawing.Printing;
using System.Media;
using System.Windows.Forms;

namespace Rich_Text_Processor
{
    public static class PublishMenuHandler
    {
        public static string[] Lines { get; set; }
        public static int LinesPrinted { get; set; }

        public static void HandlePreview(PrintPreviewDialog printPreviewDialog, PrintDocument printDocument)
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

        public static void HandlePrint(PrintDialog printDialog, PrintDocument printDocument)
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

        public static void HandlePageSetup(PageSetupDialog pageSetupDialog, PrintDocument printDocument)
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

        public static void HandleBeginPrint(MagicSpellBox magicSpellBox, PrintDialog printDialog, PrintDocument printDocument, PrintEventArgs e)
        {

            if (printDialog.PrinterSettings.PrintRange == PrintRange.Selection) Lines = magicSpellBox.SelectedText.Split(new char[] { '\n' });
            else Lines = magicSpellBox.Text.Split(new char[] { '\n' });
            int i = 0;
            foreach (string s in Lines) Lines[i++] = s.TrimEnd(new char[] { '\r' });
        }

        public static void HandlePrintPage(MagicSpellBox magicSpellBox, PrintPageEventArgs e)
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
    }
}