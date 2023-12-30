using System.Drawing;

namespace Rich_Text_Processor
{
    public interface IMagicSpellBox
    {
        string Text { get; set; }
        bool Multiline { get; set; }
        bool Modified { get; set; }
        System.Windows.Controls.RichTextBox WPFBox { get; }
        System.Windows.Forms.RichTextBox WFBox { get; }

        void SelectAll();
        void Copy();
        void Cut();
        void Paste();
        bool CanUndo { get; }
        bool CanRedo { get; }
        void Undo();
        void Redo();
        string SelectedText { get; }

        Font SelectionFont { get; set; }
        void SetSelectionColor(System.Windows.Media.Color color);
        void SetAlignment(System.Windows.TextAlignment alignment);
        void ApplySelectionForeground(System.Drawing.Color color);
    }
}