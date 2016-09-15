using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using MvvmCross.Wpf.Views;
using MyBasicTextEditor.Core.Models;

namespace MyBasicTextEditor
{
    /// <summary>
    /// Interaction logic for MyMainPage.xaml
    /// </summary>
    public partial class MyMainPage : MvxWpfView
    {
        private string docText = string.Empty;
        private Document WordDoc = new Document();

        /// <summary>
        /// Initializes a new instance of the <see cref="MyMainPage"/> class.
        /// </summary>
        public MyMainPage()
        {
            InitializeComponent();
            _fontFamily.ItemsSource = Fonts.SystemFontFamilies;
            _fontSize.ItemsSource = FontSizes;
        }

        /// <summary>
        /// Gets the font sizes.
        /// </summary>
        /// <value>
        /// The font sizes.
        /// </value>
        public double[] FontSizes
        {
            get
            {
                return new double[] { 3.0, 4.0, 5.0, 6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0, 9.5, 10.0, 10.5, 11.0, 11.5, 12.0, 12.5,13.0,13.5,14.0, 15.0,16.0, 17.0, 18.0, 19.0, 20.0, 22.0, 24.0, 26.0, 28.0, 30.0,32.0, 34.0, 36.0, 38.0, 40.0, 44.0, 48.0, 52.0, 56.0, 60.0, 64.0, 68.0, 72.0, 76.0,80.0, 88.0, 96.0, 104.0, 112.0, 120.0, 128.0, 136.0, 144.0};
            }
        }

        /// <summary>
        /// Applies the property value to selected text.
        /// </summary>
        /// <param name="DependencyPropertyformattingProperty">The dependency propertyformatting property.</param>
        /// <param name="value">The value.</param>
        void ApplyPropertyValueToSelectedText(DependencyProperty formattingProperty, object value)
        {
            if (value == null)
                return;
            Workspace.Selection.ApplyPropertyValue(formattingProperty, value);
        }

        /// <summary>
        /// Handles the SelectionChanged event of the FontFamily control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="SelectionChangedEventArgs"/> instance containing the event data.</param>
        private void FontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                FontFamily editValue = (FontFamily)e.AddedItems[0];
                ApplyPropertyValueToSelectedText(TextElement.FontFamilyProperty, editValue);
            }
            catch (Exception) { }
        }

        /// <summary>
        /// Handles the SelectionChanged event of the FontSize control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="SelectionChangedEventArgs"/> instance containing the event data.</param>
        private void FontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                ApplyPropertyValueToSelectedText(TextElement.FontSizeProperty, e.AddedItems[0]);
            }
            catch (Exception) { }
        }

        /// <summary>
        /// Updates the state of the item checked.
        /// </summary>
        /// <param name="button">The button.</param>
        /// <param name="DependencyPropertyformattingProperty">The dependency propertyformatting property.</param>
        /// <param name="expectedValue">The expected value.</param>
        void UpdateItemCheckedState(ToggleButton button, DependencyProperty formattingProperty, object expectedValue)
        {
            object currentValue = Workspace.Selection.GetPropertyValue(formattingProperty);
            button.IsChecked = (currentValue == DependencyProperty.UnsetValue) ? false : currentValue != null && currentValue.Equals(expectedValue);
        }

        /// <summary>
        /// Updates the state of the toggle button.
        /// </summary>
        private void UpdateToggleButtonState()
        {
            UpdateItemCheckedState(_btnItalic, TextElement.FontStyleProperty, FontStyles.Italic);
            UpdateItemCheckedState(_btnUnderline, Inline.TextDecorationsProperty, TextDecorations.Underline);
            UpdateItemCheckedState(_btnAlignLeft, System.Windows.Documents.Paragraph.TextAlignmentProperty, TextAlignment.Left);
            UpdateItemCheckedState(_btnAlignCenter, System.Windows.Documents.Paragraph.TextAlignmentProperty, TextAlignment.Center);
            UpdateItemCheckedState(_btnAlignRight, System.Windows.Documents.Paragraph.TextAlignmentProperty, TextAlignment.Right);
            UpdateItemCheckedState(_btnAlignJustify, System.Windows.Documents.Paragraph.TextAlignmentProperty, TextAlignment.Right);
        }

        /// <summary>
        /// Updates the type of the selection list.
        /// </summary>
        private void UpdateSelectionListType()
        {
            System.Windows.Documents.Paragraph startParagraph = Workspace.Selection.Start.Paragraph;
            System.Windows.Documents.Paragraph endParagraph = Workspace.Selection.End.Paragraph;
            System.Windows.Documents.Paragraph Paragraph = Workspace.Selection.End.Paragraph;
            if (startParagraph != null && endParagraph != null && (startParagraph.Parent is ListItem) && (endParagraph.Parent is ListItem) && object.ReferenceEquals(((ListItem)startParagraph.Parent).List, ((ListItem)endParagraph.Parent).List))
            {
                TextMarkerStyle markerStyle = ((ListItem)startParagraph.Parent).List.MarkerStyle;
                if (markerStyle == TextMarkerStyle.Disc) //bullets  
                {
                    _btnBullets.IsChecked = true;
                }
                else if (markerStyle == TextMarkerStyle.Decimal) //number  
                {
                    _btnNumbers.IsChecked = true;
                }
            }
            else
            {
                _btnBullets.IsChecked = false;
                _btnNumbers.IsChecked = false;
            }
        }

        /// <summary>
        /// Updates the selected font family.
        /// </summary>
        private void UpdateSelectedFontFamily()
        {
            object value = Workspace.Selection.GetPropertyValue(TextElement.FontFamilyProperty);
            FontFamily currentFontFamily = (FontFamily)((value == DependencyProperty.UnsetValue) ? null : value);
            if (currentFontFamily != null)
            {
                _fontFamily.SelectedItem = currentFontFamily;
            }
        }

        /// <summary>
        /// Updates the size of the selected font.
        /// </summary>
        private void UpdateSelectedFontSize()
        {
            object value = Workspace.Selection.GetPropertyValue(TextElement.FontSizeProperty);
            _fontSize.SelectedValue = (value == DependencyProperty.UnsetValue) ? null : value;
        }

        /// <summary>
        /// Updates the state of the visual.
        /// </summary>
        private void UpdateVisualState()
        {
            UpdateToggleButtonState();
            UpdateSelectionListType();
            UpdateSelectedFontFamily();
            UpdateSelectedFontSize();
        }

        /// <summary>
        /// Handles the SelectionChanged event of the Workspace control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void Workspace_SelectionChanged(object sender, RoutedEventArgs e)
        {
            UpdateVisualState();
        }

        /// <summary>
        /// Selects the img.
        /// </summary>
        /// <param name="RichTextBoxrc">The rich text boxrc.</param>
        public void selectImg(RichTextBox rc)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Image files (*.jpg, *.jpeg,*.gif, *.png) | *.jpg; *.jpeg; *.gif; *.png";
            var result = dlg.ShowDialog();
            if (result.Value)
            {
                Uri uri = new Uri(dlg.FileName, UriKind.Relative);
                BitmapImage bitmapImg = new BitmapImage(uri);
                Image image = new Image();
                image.Stretch = Stretch.Fill;
                image.Width = 250;
                image.Height = 200;
                image.Source = bitmapImg;
                var tp = rc.CaretPosition.GetInsertionPosition(LogicalDirection.Forward);
                new InlineUIContainer(image, tp);
            }
        }

        /// <summary>
        /// Handles the Click event of the btn_importimg control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void btn_importimg_Click(object sender, RoutedEventArgs e)
        {
            selectImg(Workspace);
        }

        /// <summary>
        /// Fontcolors the specified rich text boxrc.
        /// </summary>
        /// <param name="RichTextBoxrc">The rich text boxrc.</param>
        private void fontcolor(RichTextBox rc)
        {
            var colorDialog = new System.Windows.Forms.ColorDialog();
            if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var wpfcolor = Color.FromArgb(colorDialog.Color.A, colorDialog.Color.R, colorDialog.Color.G, colorDialog.Color.B);
                TextRange range = new TextRange(rc.Selection.Start, rc.Selection.End);
                range.ApplyPropertyValue(FlowDocument.ForegroundProperty, new SolidColorBrush(wpfcolor));
            }
        }

        /// <summary>
        /// Handles the Click event of the btn_Font control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void btn_Font_Click(object sender, RoutedEventArgs e)
        {
            fontcolor(Workspace);
        }

        /// <summary>
        /// Handles the Click event of the btn_SaveDoc control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void btn_SaveDoc_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog savefile = new SaveFileDialog();
            // set a default file name  
            savefile.FileName = "unknown.doc";
            // set filters - this can be done in properties as well  
            savefile.Filter = "Document files (*.doc)|*.doc";
            if (savefile.ShowDialog() == true)
            {
                TextRange t = new TextRange(Workspace.Document.ContentStart, Workspace.Document.ContentEnd);
                FileStream file = new FileStream(savefile.FileName, FileMode.Create);
                t.Save(file, System.Windows.DataFormats.Rtf);
                file.Close();
            }
        }

        /// <summary>
        /// Handles the Click event of the btn_OpenDoc control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void btn_OpenDoc_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Document files (*.doc)|*.doc";
            var result = dlg.ShowDialog();
            if (result.Value)
            {
                TextRange t = new TextRange(Workspace.Document.ContentStart, Workspace.Document.ContentEnd);
                FileStream file = new FileStream(dlg.FileName, FileMode.Open);
                t.Load(file, System.Windows.DataFormats.Rtf);
            }
        }

        /// <summary>
        /// Gets the templates.
        /// </summary>
        private void GetTemplates()
        {
            DirectoryInfo dir = new DirectoryInfo(@"C:\Users\andrew.rae\Desktop\TestTemplates\");
            var templates = dir.GetFiles("*.dotx", SearchOption.AllDirectories);

            if (templates != null)
            {
                List<string> templateNames = new List<string>();

                foreach (var template in templates)
                {
                    templateNames.Add(template.Name.Substring(0, template.Name.IndexOf('.')));
                }
                ((MyMainPageViewModel)this.ViewModel).SetTemplates(templateNames);
            }
        }

        /// <summary>
        /// Handles the Click event of the SaveBttn control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void SaveBttn_Click(object sender, RoutedEventArgs e)
        {
            object missing = System.Reflection.Missing.Value;
            object Visible = true;
            object start1 = 0;
            object end1 = 0;

            ApplicationClass WordApp = new ApplicationClass();
            Document adoc = WordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            Range rng = adoc.Range(ref start1, ref missing);

            try
            {
                rng.Font.Name = "Georgia";
                rng.InsertAfter(new TextRange(Workspace.Document.ContentStart, Workspace.Document.ContentEnd).Text);
                object filename = @"C:\Users\andrew.rae\Desktop\MyWord.doc";
                adoc.SaveAs(ref filename);
                WordApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Handles the Click event of the replaceTagBttn control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void replaceTagBttn_Click(object sender, RoutedEventArgs e)
        {
            var currentViewModel = this.ViewModel as MyMainPageViewModel;

            if (currentViewModel.SelectedPatient != null)
            {
                string newText = currentViewModel.ReplaceAllTags(this.docText);

                Workspace.Document.Blocks.Clear();
                this.Insert(newText);
            }
            else
            {
                MessageBox.Show("please select the patient whos data you want to replace the tags with", "Select Tag");
            }

        }

        /// <summary>
        /// Handles the Click event of the LaunchBttn control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void LaunchBttn_Click(object sender, RoutedEventArgs e)
        {
            ApplicationClass WordApp = new ApplicationClass();
            var currentViewModel = this.ViewModel as MyMainPageViewModel;
            if (currentViewModel.SelectedTemplate != null)
            {
                object templatePath = @"C:\Users\andrew.rae\Desktop\TestTemplates\" + currentViewModel.SelectedTemplate + ".dotx";
                string saveFilePath = @"C:\Users\andrew.rae\Desktop\TestTemplates\";

                WordDoc = WordApp.Documents.Add(templatePath); // open the template
                WordDoc.SaveAs(saveFilePath + "TestBPLetter" + currentViewModel.SelectedPatient.Forename); // save template as document so you ont overwirte the template
                WordDoc = WordApp.Documents.Open(saveFilePath + "TestBPLetter" + currentViewModel.SelectedPatient.Forename + ".docx"); // open the newly saved doc to read the text

                Range rng = WordDoc.Range();
                int count = WordDoc.Sections.Count;
                Range headerRange = WordDoc.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                Range footerRange = WordDoc.Sections[count].Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;


                Find documentFindObject = rng.Find;
                Find headerFindObject = headerRange.Find;
                Find footerFindObject = footerRange.Find;

                object replaceAll = WdReplace.wdReplaceAll;

                foreach (Tags tag in currentViewModel.TagList)
                {
                    //// Clear search box
                    documentFindObject.ClearFormatting();
                    headerFindObject.ClearFormatting();
                    footerFindObject.ClearFormatting();

                    //// Search for tag
                    documentFindObject.Text = tag.Tag;
                    headerFindObject.Text = tag.Tag;
                    footerFindObject.Text = tag.Tag;

                    //// Clear search tags
                    documentFindObject.Replacement.ClearFormatting();
                    headerFindObject.Replacement.ClearFormatting();
                    footerFindObject.Replacement.ClearFormatting();

                    //// Set up text to replace with
                    documentFindObject.Replacement.Text = currentViewModel.ReplaceTag(tag.Tag);
                    headerFindObject.Replacement.Text = currentViewModel.ReplaceTag(tag.Tag);
                    footerFindObject.Replacement.Text = currentViewModel.ReplaceTag(tag.Tag);

                    //// Run replacement
                    documentFindObject.Execute(Replace: replaceAll);
                    headerFindObject.Execute(Replace: replaceAll);
                    footerFindObject.Execute(Replace: replaceAll);
                }
            }
            else
            {
                MessageBox.Show("please select a template from the drop down list", "select a template");
            }
        }

        /// <summary>
        /// Handles the Click event of the TemplateBttn control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void TemplateBttn_Click(object sender, RoutedEventArgs e)
        {
            this.GetTemplates();
        }

        /// <summary>
        /// Handles the Click event of the printWithOptions control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void printWithOptions_Click(object sender, RoutedEventArgs e)
        {
            ApplicationClass WordApp = new ApplicationClass();

            
            WordDoc = WordApp.Documents.Add();
            WordApp.ActiveDocument.PageSetup.TopMargin = 20;
            WordApp.ActiveDocument.PageSetup.BottomMargin = 20;
            WordApp.ActiveDocument.PageSetup.RightMargin = 15;
            WordApp.ActiveDocument.PageSetup.LeftMargin = 15;

            this.FormatWordDoc();

            Dialog dialog = WordApp.Dialogs[WdWordDialog.wdDialogFilePrint];
            var dialogResult = dialog.Show();
        }

        /// <summary>
        /// Handles the Click event of the SilentHiddenPrint control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void SilentHiddenPrint_Click(object sender, RoutedEventArgs e)
        {
            ApplicationClass WordApp = new ApplicationClass();

            WordApp.Visible = true;
            WordDoc = WordApp.Documents.Add();
            WordApp.ActiveDocument.PageSetup.TopMargin = 40;
            WordApp.ActiveDocument.PageSetup.BottomMargin = 40;
            WordApp.ActiveDocument.PageSetup.RightMargin = 35;
            WordApp.ActiveDocument.PageSetup.LeftMargin = 35;

            this.FormatWordDoc();

            ///WordDoc.PrintOut();
        }

        /// <summary>
        /// Formats the word document.
        /// </summary>
        private void FormatWordDoc()
        {
            int characterCount = 0;
            foreach (System.Windows.Documents.Paragraph para in Workspace.Document.Blocks)
            {
                var wordPara = WordDoc.Content.Paragraphs.Add();

                var paragraphText = new TextRange(para.ContentStart, para.ContentEnd).Text;
                wordPara.Range.Text = paragraphText;

                foreach (var inline in para.Inlines)
                {
                    var inlineText = new TextRange(inline.ContentStart, inline.ContentEnd);
                    var rangeStart = WordDoc.Paragraphs.Last.Range.Text.IndexOf(inlineText.Text) + characterCount;
                    Range formatRange = WordDoc.Range(rangeStart, rangeStart + inlineText.Text.Length);

                    formatRange.Font.Underline = inlineText.GetPropertyValue(Inline.TextDecorationsProperty) == TextDecorations.Underline ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone;
                    formatRange.Font.Italic = inlineText.GetPropertyValue(TextElement.FontStyleProperty).ToString().Equals("Italic") ? 1 : 0;
                    formatRange.Font.Bold = inlineText.GetPropertyValue(TextElement.FontWeightProperty).ToString().Equals("Bold") ? 1 : 0;

                    SolidColorBrush SCBrush = (SolidColorBrush)inline.Foreground;
                    formatRange.Font.Color = (WdColor)(SCBrush.Color.R + 0x100 * SCBrush.Color.G + 0x10000 * SCBrush.Color.B);
                    formatRange.Font.Size = (float)inline.FontSize;
                    formatRange.Font.Name = inline.FontFamily.ToString();
                    
                    formatRange.ParagraphFormat.Alignment =
                        inlineText.GetPropertyValue(System.Windows.Documents.Paragraph.TextAlignmentProperty).ToString().Equals("Center") ? WdParagraphAlignment.wdAlignParagraphCenter :
                        inlineText.GetPropertyValue(System.Windows.Documents.Paragraph.TextAlignmentProperty).ToString().Equals("Right") ? WdParagraphAlignment.wdAlignParagraphRight :
                        inlineText.GetPropertyValue(System.Windows.Documents.Paragraph.TextAlignmentProperty).ToString().Equals("Justify") ? WdParagraphAlignment.wdAlignParagraphJustify :
                        WdParagraphAlignment.wdAlignParagraphLeft;
                }
                characterCount += WordDoc.Paragraphs.Last.Range.Text.Length;
                wordPara.Range.InsertParagraphAfter();
                
            }
        }

        /// <summary>
        /// Handles the TextChanged event of the rtbEditor control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.Controls.TextChangedEventArgs"/> instance containing the event data.</param>
        private void rtbEditor_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            docText = new TextRange(this.Workspace.Document.ContentStart, this.Workspace.Document.ContentEnd).Text;
        }

        /// <summary>
        /// Handles the Click event of the insertTagBttn control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void insertTagBttn_Click(object sender, RoutedEventArgs e)
        {
            var currentViewModel = this.ViewModel as MyMainPageViewModel;

            if (currentViewModel.SelectedTag != null)
            {
                this.Insert(currentViewModel.SelectedTag.Tag);
            }
            else
            {
                MessageBox.Show("please select a tag to be inserted", "Select Tag");
            }
        }

        /// <summary>
        /// Inserts the specified insert string.
        /// </summary>
        /// <param name="insertString">The insert string.</param>
        private void Insert(string insertString)
        {
            Workspace.CaretPosition.InsertTextInRun(insertString);

            TextPointer moveTo = Workspace.CaretPosition.GetNextContextPosition(LogicalDirection.Forward);

            if (moveTo != null)
            {
                Workspace.CaretPosition = moveTo;
            }
        }
    }
}