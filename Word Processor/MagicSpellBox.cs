﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design.Serialization;
using System.Drawing;
using System.Linq;
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
        public MagicSpellBox()
        {
            Box = new RichTextBox { IsReadOnly = false, SpellCheck = { IsEnabled = true }, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            Box.TextChanged += (s, e) => MagicSpellBox_TextChanged(s, e);
            base.Child = Box;
            Font = new Font("Microsoft Sans Serif", 10, FontStyle.Regular);
            Multiline = true;
        }

        #region Properties

        public RichTextBox Box { get; }

        [Browsable(true)]
        [Category("Extended Properties")]
        [Description("If the RichTextBox's Text has been changed")]
        public bool Modified { get; set; }

        [Browsable(true)]
        [Category("Action")]
        [Description("Invoked when Text Changes")]
        public new event EventHandler TextChanged;

        [DefaultValue(false)]
        public override string Text
        {
            get => new TextRange(Box.Document.ContentStart, Box.Document.ContentEnd).Text;
            set
            {
                Box.Document.Blocks.Clear();
                Box.Document.Blocks.Add(new Paragraph(new Run(value)));
            }
        }

        [DefaultValue(false)]
        public bool Multiline
        {
            get => Box.AcceptsReturn;
            set => Box.AcceptsReturn = value;
        }

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public new UIElement Child
        {
            get => base.Child;
            set { /* Do nothing */ } // This must be done to fix an issue with the serializer
        }

        public Font SelectionFont
        {
            get => ConvertToFont(Box.Selection.GetPropertyValue(Control.FontFamilyProperty), Box.Selection.GetPropertyValue(Control.FontSizeProperty));
            set => ApplyFont(value);
        }

        public bool CanUndo => Box.CanUndo;
        public bool CanRedo => Box.CanRedo;
        public string SelectedText => Box.Selection.Text;
        public int WordCount => Text.Trim().Split(new char[] { ' ', '\t', '\r', '\n', '\v', '\f' }, StringSplitOptions.RemoveEmptyEntries).Length;
        public int CharCount => new TextRange(Box.Document.ContentStart, Box.Document.ContentEnd).Text.Count(c => !char.IsWhiteSpace(c));

        #endregion

        #region Events

        protected void MagicSpellBox_TextChanged(object sender, EventArgs e)
        {
            TextChanged?.Invoke(this, e);
            Modified = true;
        }

        #endregion

        #region Methods

        public void Undo() => Box.Undo();
        public void Redo() => Box.Redo();
        public void SelectAll() => Box.SelectAll();
        public void Copy() => Box.Copy();
        public void Cut() => Box.Cut();
        public void Paste() => Box.Paste();
        private Font ConvertToFont(object fontFamily, object fontSize) => new Font(fontFamily as string, (float)(double)Convert.ToDouble(fontSize));

        private void ApplyFont(Font font)
        {
            if (font != null)
            {
                Box.Selection.Start.Paragraph.FontFamily = new System.Windows.Media.FontFamily(font.FontFamily.Name);                                                                                     // Apply font family and size
                Box.Selection.Start.Paragraph.FontSize = font.Size;
                if (font.Underline) new TextRange(Box.Selection.Start, Box.Selection.End).ApplyPropertyValue(Inline.TextDecorationsProperty, new TextDecorationCollection { TextDecorations.Underline }); // Simulate underline using TextDecorations
                else new TextRange(Box.Selection.Start, Box.Selection.End).ApplyPropertyValue(Inline.TextDecorationsProperty, null);                                                                      // Clear underline
            }
        }

        public void SetSelectionColor(System.Windows.Media.Color color)
        {
            if (Box.Selection != null)
                if (Box.Selection.Text.Length > 0) foreach (Run run in GetRunsInSelection()) run.Foreground = new SolidColorBrush(color); // Apply color to each Run element in the selection
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

        public void Bold()
        {
            if (Box.Selection != null && !Box.Selection.IsEmpty)
            {
                var currentWeight = Box.Selection.GetPropertyValue(TextElement.FontWeightProperty);
                Box.Selection.ApplyPropertyValue(TextElement.FontWeightProperty, (currentWeight != null && currentWeight.Equals(FontWeights.Bold)) ? FontWeights.Normal : FontWeights.Bold);
            }
        }

        public void Italic()
        {
            if (Box.Selection != null && !Box.Selection.IsEmpty)
            {
                var currentStyle = Box.Selection.GetPropertyValue(TextElement.FontStyleProperty);
                Box.Selection.ApplyPropertyValue(TextElement.FontStyleProperty, (currentStyle != null && currentStyle.Equals(FontStyles.Italic)) ? FontStyles.Normal : FontStyles.Italic);
            }
        }

        public void Underline()
        {
            if (SelectionFont != null) new TextRange(Box.Selection.Start, Box.Selection.End).ApplyPropertyValue(Inline.TextDecorationsProperty, TextDecorations.Underline);
        }

        public void Bullet()
        {
            if (!new TextRange(Box.Selection.Start, Box.Selection.End).IsEmpty)                            // Check if the selection is empty
            {
                string[] lines = new TextRange(Box.Selection.Start, Box.Selection.End).Text.Split('\n');   // Split the selected text into lines
                new TextRange(Box.Selection.Start, Box.Selection.End).Text = "";                           // Clear the existing selection

                foreach (string line in lines)                                                             // Insert a bullet point at the beginning of each line
                {
                    if (!string.IsNullOrWhiteSpace(line))                                                  // Check if the line is not empty
                    {
                        Box.CaretPosition.InsertTextInRun($"\u2022 {line.Trim()}");
                        Box.CaretPosition = Box.CaretPosition.GetPositionAtOffset(4 + line.Trim().Length); // Move the caret to the end of the inserted text
                    }

                    Box.CaretPosition.InsertParagraphBreak();                                              // Insert a newline after each line
                }
            }
        }

        public void SetAlignment(TextAlignment alignment)
        {
            if (Box.Selection.Start.Paragraph != null) Box.Selection.Start.Paragraph.TextAlignment = alignment;
        }

        public void ApplySelectionForeground(System.Drawing.Color color)
        {
            if (Box.Selection != null) new TextRange(Box.Selection.Start, Box.Selection.End).ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(color.ToMediaColor()));
        }

        #endregion
    }
}