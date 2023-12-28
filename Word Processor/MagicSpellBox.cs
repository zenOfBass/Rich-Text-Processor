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
            base.Child = box;
            box.IsReadOnly = false;
            box.TextChanged += (s, e) => OnTextChanged(EventArgs.Empty);
            box.SpellCheck.IsEnabled = true;
            box.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            Font = new Font("Microsoft Sans Serif", 10, FontStyle.Regular);
            Multiline = true;
            Size = new System.Drawing.Size(100, 20);
        }

        [DefaultValue(false)]
        public override string Text
        {
            get
            {
                string richText = new TextRange(box.Document.ContentStart, box.Document.ContentEnd).Text;
                return richText;
            }
            set
            {
                box.Document.Blocks.Clear();
                box.Document.Blocks.Add(new Paragraph(new Run(value)));
            }
        }

        [DefaultValue(false)]
        public bool Multiline
        {
            get { return box.AcceptsReturn; }
            set { box.AcceptsReturn = value; }
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

        public void SelectAll() => box.SelectAll();

        public void Copy() => box.Copy();

        public void Cut() => box.Cut();

        public void Paste() => box.Paste();

        public bool CanUndo => box.CanUndo;

        public bool CanRedo => box.CanRedo;

        public void Undo() => box.Undo();

        public void Redo() => box.Redo();

        public string SelectedText => box.Selection.Text;

        private Font ConvertToFont(object fontFamily, object fontSize)
        {
            string familyName = fontFamily as string;
            double size = Convert.ToDouble(fontSize);

            return new Font(familyName, (float)size);
        }

        public Font SelectionFont
        {
            get { return ConvertToFont(box.Selection.GetPropertyValue(Control.FontFamilyProperty), box.Selection.GetPropertyValue(Control.FontSizeProperty)); }
            set { ApplyFont(value); }
        }

        private void ApplyFont(Font font)
        {
            if (font != null)
            {
                TextPointer start = box.Selection.Start;
                TextPointer end = box.Selection.End;

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
            if (box.Selection != null)
            {
                // Check if the selection is actually text
                if (box.Selection.Text.Length > 0)
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
            var textPointer = box.Selection.Start;
            var runs = new List<Run>();

            while (textPointer.CompareTo(box.Selection.End) < 0)
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
            var paragraph = box.Selection.Start.Paragraph;
            if (paragraph != null) paragraph.TextAlignment = alignment;
        }

        public void ApplySelectionForeground(System.Drawing.Color color)
        {
            if (box.Selection != null)
            {
                var selectionRange = new TextRange(box.Selection.Start, box.Selection.End);
                selectionRange.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(color.ToMediaColor()));
            }
        }
    }
}