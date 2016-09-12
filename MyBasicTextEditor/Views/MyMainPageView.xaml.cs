using Microsoft.Office.Interop.Word;
using MvvmCross.Wpf.Views;
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Documents;
using MyBasicTextEditor.Core.Models;
using System.IO;

namespace MyBasicTextEditor
{
    /// <summary>
    /// Interaction logic for MyMainPage.xaml
    /// </summary>
    public partial class MyMainPage : MvxWpfView
    {
        public MyMainPage()
        {
            InitializeComponent();
        }

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
                rng.InsertAfter(new TextRange(rtbEditor.Document.ContentStart, rtbEditor.Document.ContentEnd).Text);
                object filename = @"C:\Users\andrew.rae\Desktop\MyWord.doc";
                adoc.SaveAs(ref filename);
                WordApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void replaceTagBttn_Click(object sender, RoutedEventArgs e)
        {
            var currentViewModel = this.ViewModel as MyMainPageViewModel;

        }

        private void LaunchBttn_Click(object sender, RoutedEventArgs e)
        {
            ApplicationClass WordApp = new ApplicationClass();
            Document wordDoc = new Document();
            var currentViewModel = this.ViewModel as MyMainPageViewModel;
            if (currentViewModel.SelectedTemplate != null)
            {
                object templatePath = @"C:\Users\andrew.rae\Desktop\TestTemplates\" + currentViewModel.SelectedTemplate + ".dotx";
                string saveFilePath = @"C:\Users\andrew.rae\Desktop\TestTemplates\";

                wordDoc = WordApp.Documents.Add(templatePath); // open the template
                wordDoc.SaveAs(saveFilePath + "TestBPLetter" + currentViewModel.SelectedPatient.Forename); // save template as document so you ont overwirte the template
                wordDoc = WordApp.Documents.Open(saveFilePath + "TestBPLetter" + currentViewModel.SelectedPatient.Forename + ".docx"); // open the newly saved doc to read the text

                Range rng = wordDoc.Range();
                int count = wordDoc.Sections.Count;
                Range headerRange = wordDoc.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                Range footerRange = wordDoc.Sections[count].Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;


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

        private void TemplateBttn_Click(object sender, RoutedEventArgs e)
        {
            this.GetTemplates();
        }
    }
}