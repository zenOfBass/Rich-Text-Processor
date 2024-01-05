using System.Windows.Forms;

namespace Rich_Text_Processor
{
    public static class WordAndCharCountHandler
    {
        public static void HandleWordCount(MagicSpellBox magicSpellBox, ToolStripStatusLabel labelWordCount) => labelWordCount.Text = string.IsNullOrEmpty(magicSpellBox.Text.Trim())
                ? "0 words" : magicSpellBox.WordCount <= 1 ? "1 word" : $"{magicSpellBox.WordCount} words";

        public static void HandleCharCount(MagicSpellBox magicSpellBox, ToolStripStatusLabel labelCharCount) => labelCharCount.Text = $"{magicSpellBox.CharCount} characters";
    }
}