using Rich_Text_Processor;
using System;
using System.Media;
using System.Windows.Forms;

namespace Word_Processor
{
    public static class EditMenuHandler
    {
        public static void HandleSelectAll(MagicSpellBox magicSpellBox)
        {
            try
            {
                magicSpellBox.SelectAll();
            }
            catch (Exception)
            {
                ShowErrorMessage("Unable to select all document content.", "RTE - Select");
            }
        }

        public static void HandleCopy(MagicSpellBox magicSpellBox)
        {
            try
            {
                magicSpellBox.Copy();
            }
            catch (Exception)
            {
                ShowErrorMessage("Unable to copy document content.", "RTE - Copy");
            }
        }

        public static void HandleCut(MagicSpellBox magicSpellBox)
        {
            try
            {
                magicSpellBox.Cut();
            }
            catch
            {
                ShowErrorMessage("Unable to cut document content.", "RTE - Cut");
            }
        }

        public static void HandlePaste(MagicSpellBox magicSpellBox)
        {
            try
            {
                magicSpellBox.Paste();
            }
            catch
            {
                ShowErrorMessage("Unable to copy clipboard content to document.", "RTE - Paste");
            }
        }

        public static void HandleUndo(MagicSpellBox magicSpellBox)
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

        public static void HandleRedo(MagicSpellBox magicSpellBox)
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

        private static void ShowErrorMessage(string message, string caption) => MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
