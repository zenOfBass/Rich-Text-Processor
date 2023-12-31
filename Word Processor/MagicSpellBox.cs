using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design.Serialization;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Forms.Design;
using System.Windows.Forms.Integration;
using System.Windows.Media;
using FontStyle = System.Drawing.FontStyle;

namespace Rich_Text_Processor
{
    [Designer(typeof(ControlDesigner))]
    [DesignerSerializer("System.Windows.Forms.Design.ControlCodeDomSerializer, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a",
                "System.ComponentModel.Design.Serialization.CodeDomSerializer, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
    public class MagicSpellBox : ElementHost
    {
        private readonly RichTextBox box;

        public MagicSpellBox()
        {
            box = new RichTextBox();
            Box.IsReadOnly = false;
            Box.TextChanged += (s, e) => OnTextChanged(EventArgs.Empty);
            Box.SpellCheck.IsEnabled = true;
            Box.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            base.Child = Box;
            Font = new Font("Microsoft Sans Serif", 10, FontStyle.Regular);
            Multiline = true;
            Size = new System.Drawing.Size(100, 20);
        }

        [DefaultValue(false)]
        public override string Text
        {
            get
            {
                string richText = new TextRange(Box.Document.ContentStart, Box.Document.ContentEnd).Text;
                return richText;
            }
            set
            {
                Box.Document.Blocks.Clear();
                Box.Document.Blocks.Add(new Paragraph(new Run(value)));
            }
        }

        [DefaultValue(false)]
        public bool Multiline
        {
            get { return Box.AcceptsReturn; }
            set { Box.AcceptsReturn = value; }
        }

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public new UIElement Child
        {
            get { return base.Child; }
            set { /* Do nothing */ }
        }

        [Browsable(true)]
        [Category("Extended Properties")]
        [Description("Set TextBox border Color")]
        public bool Modified
        {
            get;
            set;
        }

        public RichTextBox Box => box;

        public void SelectAll() => Box.SelectAll();

        public void Copy() => Box.Copy();

        public void Cut() => Box.Cut();

        public void Paste() => Box.Paste();

        public bool CanUndo => Box.CanUndo;

        public bool CanRedo => Box.CanRedo;

        public void Undo() => Box.Undo();

        public void Redo() => Box.Redo();

        public string SelectedText => Box.Selection.Text;

        private Font ConvertToFont(object fontFamily, object fontSize)
        {
            string familyName = fontFamily as string;
            double size = Convert.ToDouble(fontSize);

            return new Font(familyName, (float)size);
        }

        public Font SelectionFont
        {
            get { return ConvertToFont(Box.Selection.GetPropertyValue(Control.FontFamilyProperty), Box.Selection.GetPropertyValue(Control.FontSizeProperty)); }
            set { ApplyFont(value); }
        }

        private void ApplyFont(Font font)
        {
            if (font != null)
            {
                TextPointer start = Box.Selection.Start;
                TextPointer end = Box.Selection.End;

                // Apply font family and size
                start.Paragraph.FontFamily = new System.Windows.Media.FontFamily(font.FontFamily.Name);
                start.Paragraph.FontSize = font.Size;

                // Simulate underline using TextDecorations
                if (font.Underline)
                {
                    TextDecorationCollection decorations = new TextDecorationCollection();
                    decorations.Add(TextDecorations.Underline);
                    TextRange textRange = new TextRange(start, end);
                    textRange.ApplyPropertyValue(Inline.TextDecorationsProperty, decorations);
                }
                else
                {
                    // Clear underline
                    TextRange textRange = new TextRange(start, end);
                    textRange.ApplyPropertyValue(Inline.TextDecorationsProperty, null);
                }
            }
        }

        public void SetSelectionColor(System.Windows.Media.Color color)
        {
            if (Box.Selection != null)
            {
                // Check if the selection is actually text
                if (Box.Selection.Text.Length > 0)
                {
                    // Apply color to each Run element in the selection
                    foreach (Run run in GetRunsInSelection())
                    {
                        run.Foreground = new SolidColorBrush(color);
                    }
                }
            }
        }

        private IEnumerable<Run> GetRunsInSelection()
        {
            var textPointer = Box.Selection.Start;
            var runs = new List<Run>();

            while (textPointer.CompareTo(Box.Selection.End) < 0)
            {
                if (textPointer.Parent is Run run)
                {
                    runs.Add(run);
                    textPointer = run.ElementEnd.GetNextInsertionPosition(LogicalDirection.Forward);
                }
            }

            return runs;
        }

        public void SetAlignment(TextAlignment alignment)
        {
            var paragraph = Box.Selection.Start.Paragraph;
            if (paragraph != null) paragraph.TextAlignment = alignment;
        }

        public void ApplySelectionForeground(System.Drawing.Color color)
        {
            if (Box.Selection != null)
            {
                var selectionRange = new TextRange(Box.Selection.Start, Box.Selection.End);
                selectionRange.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(color.ToMediaColor()));
            }
        }
    }
}