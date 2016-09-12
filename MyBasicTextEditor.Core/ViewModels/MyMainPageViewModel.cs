using MvvmCross.Core.ViewModels;
using MyBasicTextEditor.Core.Models;
using System;
using System.Collections.Generic;

namespace MyBasicTextEditor
{
    public class MyMainPageViewModel : MvxViewModel
    {
        private List<Patient> patientList;
        private Patient selectedPatient;

        private List<Tags> tagList;

        private List<string> templateList;
        private string selectedTemplate;
        public MyMainPageViewModel()
        {
            this.PatientList = new List<Patient>()
            {
                new Patient("Jason", "Borne", DateTime.Parse("18/02/1962"),123456789),
                new Patient("James", "Kirk", DateTime.Parse("20/06/1950"),123456788,new List<string> () {"Tiberius"}),
                new Patient("Ethan", "Hunt", DateTime.Parse("05/10/1996"),123456777),
                new Patient("Alex", "Rider", DateTime.Parse("01/01/2006"),123456666, new List<string>() {"Dave"}),
                new Patient("Sherlock", "Holmes", DateTime.Parse("01/06/1854"),123455555),
                new Patient("Tony", "Stark", DateTime.Parse("29/05/1970"),123444444),
                new Patient("Albus", "Dumbledore", DateTime.Parse("29/05/1881"),123444444, new List<string>() {"Percival", "Wulfric", "Brian"})
            };

            this.TagList = new List<Tags>()
            {
                new Tags("Forename",Tags.Forename),
                new Tags("Surname",Tags.Surname),
                new Tags("Middle names",Tags.Middlenames),
                new Tags("Full name",Tags.Fullname),
                new Tags("Display name",Tags.Displayname),
                new Tags("Date of birth",Tags.Dateofbirth),
                new Tags("Id number",Tags.Idnumber),
                new Tags("Initialled name",Tags.InitialledName),
                new Tags("Main address",Tags.MainAddress),
            };

            this.SelectedPatient = PatientList[0];
        }
        

        public List<Patient> PatientList
        {
            get { return patientList; }
            set { patientList = value; }
        }

        public Patient SelectedPatient
        {
            get { return selectedPatient; }
            set { selectedPatient = value; }
        }

        public List<Tags> TagList
        {
            get { return this.tagList; }
            set { this.tagList = value; }
        }

        public List<string> TemplateList
        {
            get { return this.templateList; }
            set { this.SetProperty(ref this.templateList, value); }
        }

        public string SelectedTemplate
        {
            get { return this.selectedTemplate; }
            set { this.SetProperty(ref this.selectedTemplate, value); }
        }

        public string ReplaceTag(string tag)
        {
            switch (tag)
            {
                case Tags.Forename:
                    {
                        return SelectedPatient.Forename;
                    }
                case Tags.Surname:
                    {
                        return SelectedPatient.Surname;
                    }
                case Tags.Fullname:
                    {
                        return SelectedPatient.FullName;
                    }
                case Tags.Displayname:
                    {
                        return SelectedPatient.DisplayName;
                    }
                case Tags.Dateofbirth:
                    {
                        return SelectedPatient.DOB.ToString();
                    }
                case Tags.Idnumber:
                    {
                        return SelectedPatient.PatientNumber.ToString();
                    }
                case Tags.MainAddress:
                    {
                        return SelectedPatient.PrimaryAddress.DisplayAddress;
                    }
                case Tags.InitialledName:
                    {
                        return SelectedPatient.InitialledName;
                    }
                default:
                    {
                        return "Error Tag Not Found";
                    }

            }
        }

        public void SetTemplates(List<string> templates)
        {
            this.TemplateList = templates;
            this.SelectedTemplate = this.TemplateList[0];
        }
    }
}
