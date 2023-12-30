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
    public class MagicSpellBox : ElementHost, IMagicSpellBox
    {
        public MagicSpellBox()
        {
            WPFBox = new RichTextBox();
            base.Child = WPFBox;
            WPFBox.IsReadOnly = false;
            WPFBox.TextChanged += (s, e) => OnTextChanged(EventArgs.Empty);
            WPFBox.SpellCheck.IsEnabled = true;
            WPFBox.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            Font = new Font("Microsoft Sans Serif", 10, FontStyle.Regular);
            Multiline = true;
            Size = new System.Drawing.Size(100, 20);
        }

        [DefaultValue(false)]
        public override string Text
        {
            get
            {
                string richText = new TextRange(WPFBox.Document.ContentStart, WPFBox.Document.ContentEnd).Text;
                return richText;
            }
            set
            {
                WPFBox.Document.Blocks.Clear();
                WPFBox.Document.Blocks.Add(new Paragraph(new Run(value)));
            }
        }

        [DefaultValue(false)]
        public bool Multiline
        {
            get { return WPFBox.AcceptsReturn; }
            set { WPFBox.AcceptsReturn = value; }
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

        public System.Windows.Controls.RichTextBox WPFBox { get; }

        public System.Windows.Forms.RichTextBox WFBox { get; }

        public void SelectAll() => WPFBox.SelectAll();

        public void Copy() => WPFBox.Copy();

        public void Cut() => WPFBox.Cut();

        public void Paste() => WPFBox.Paste();

        public bool CanUndo => WPFBox.CanUndo;

        public bool CanRedo => WPFBox.CanRedo;

        public void Undo() => WPFBox.Undo();

        public void Redo() => WPFBox.Redo();

        public string SelectedText => WPFBox.Selection.Text;

        private Font ConvertToFont(object fontFamily, object fontSize)
        {
            string familyName = fontFamily as string;
            double size = Convert.ToDouble(fontSize);

            return new Font(familyName, (float)size);
        }

        public Font SelectionFont
        {
            get { return ConvertToFont(WPFBox.Selection.GetPropertyValue(Control.FontFamilyProperty), WPFBox.Selection.GetPropertyValue(Control.FontSizeProperty)); }
            set { ApplyFont(value); }
        }

        private void ApplyFont(Font font)
        {
            if (font != null)
            {
                TextPointer start = WPFBox.Selection.Start;
                TextPointer end = WPFBox.Selection.End;

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
            if (WPFBox.Selection != null)
            {
                // Check if the selection is actually text
                if (WPFBox.Selection.Text.Length > 0)
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
            var textPointer = WPFBox.Selection.Start;
            var runs = new List<Run>();

            while (textPointer.CompareTo(WPFBox.Selection.End) < 0)
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
            var paragraph = WPFBox.Selection.Start.Paragraph;
            if (paragraph != null) paragraph.TextAlignment = alignment;
        }

        public void ApplySelectionForeground(System.Drawing.Color color)
        {
            if (WPFBox.Selection != null)
            {
                var selectionRange = new TextRange(WPFBox.Selection.Start, WPFBox.Selection.End);
                selectionRange.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(color.ToMediaColor()));
            }
        }
    }
}