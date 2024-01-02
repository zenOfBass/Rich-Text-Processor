using System.Media;
using System.Windows;
using System.Windows.Forms;

namespace Rich_Text_Processor
{
    public static class FormatToolbarHandler
    {
        public static void HandleFontSelect(MagicSpellBox magicSpellBox, FontDialog fontDialog)
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

        public static void HandleFontColor(MagicSpellBox magicSpellBox, ColorDialog colorDialog)
        {
            try
            {
                colorDialog.Color = magicSpellBox.ForeColor;
                if (colorDialog.ShowDialog() == DialogResult.OK) magicSpellBox.ApplySelectionForeground(colorDialog.Color);
            }
            catch
            {
                SystemSounds.Hand.Play();
            }
        }

        public static void HandleBold(MagicSpellBox magicSpellBox)
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

        public static void HandleItalic(MagicSpellBox magicSpellBox)
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

        public static void HandleUnderline(MagicSpellBox magicSpellBox)
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

        public static void HandleAlignLeft(MagicSpellBox magicSpellBox)
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

        public static void HandleAlignCenter(MagicSpellBox magicSpellBox)
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

        public static void HandleAlignRight(MagicSpellBox magicSpellBox)
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

        public static void HandleBullets(MagicSpellBox magicSpellBox)
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
    }
}