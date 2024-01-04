using System;
using System.Media;

namespace Rich_Text_Processor
{
    public static class EditMenuHandler
    {
        public static void HandleSelectAll(MagicSpellBox magicSpellBox)
        {
            try
            {
                magicSpellBox.SelectAll();
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling Select All: {ex.Message}");
                SystemSounds.Hand.Play();
            }
        }

        public static void HandleCopy(MagicSpellBox magicSpellBox)
        {
            try
            {
                magicSpellBox.Copy();
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling Copy: {ex.Message}");
                SystemSounds.Hand.Play();
            }
        }

        public static void HandleCut(MagicSpellBox magicSpellBox)
        {
            try
            {
                magicSpellBox.Cut();
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling Cut: {ex.Message}");
                SystemSounds.Hand.Play();
            }
        }

        public static void HandlePaste(MagicSpellBox magicSpellBox)
        {
            try
            {
                magicSpellBox.Paste();
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling Paste: {ex.Message}");
                SystemSounds.Hand.Play();
            }
        }

        public static void HandleUndo(MagicSpellBox magicSpellBox)
        {
            try
            {
                if (magicSpellBox.CanUndo) magicSpellBox.Undo();
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling Undo: {ex.Message}");
                SystemSounds.Hand.Play();
            }
        }

        public static void HandleRedo(MagicSpellBox magicSpellBox)
        {
            try
            {
                if (magicSpellBox.CanRedo) magicSpellBox.Redo();
            }
            catch (Exception ex)
            {
                Logger.Log(LogLevel.Error, $"Error handling Redo: {ex.Message}");
                SystemSounds.Hand.Play();
            }
        }
    }
}