﻿using Microsoft.Office.Interop.Word;
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
            DirectoryInfo dir = new DirectoryInfo(@"C:\Users\conor.nicol\Desktop\TestTemplates\");
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
                object filename = @"C:\Users\conor.nicol\Desktop\MyWord.doc";
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
                object templatePath = @"C:\Users\conor.nicol\Desktop\TestTemplates\" + currentViewModel.SelectedTemplate + ".dotx";
                string saveFilePath = @"C:\Users\conor.nicol\Desktop\TestTemplates\";

                wordDoc = WordApp.Documents.Add(templatePath); // open the template
                wordDoc.SaveAs(saveFilePath + "TestBPLetter" + currentViewModel.SelectedPatient.Forename); // save template as document so you ont overwirte the template
                wordDoc = WordApp.Documents.Open(saveFilePath + "TestBPLetter" + currentViewModel.SelectedPatient.Forename + ".docx"); // open the newly saved doc to read the text

                Range rng = wordDoc.Range();

                Find findObject = rng.Find;
                object replaceAll = WdReplace.wdReplaceAll;

                foreach (Tags tag in currentViewModel.TagList)
                {
                    findObject.ClearFormatting();
                    findObject.Text = tag.Tag;
                    findObject.Replacement.ClearFormatting();
                    findObject.Replacement.Text = currentViewModel.ReplaceTag(tag.Tag);
                    findObject.Execute(Replace: replaceAll);
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