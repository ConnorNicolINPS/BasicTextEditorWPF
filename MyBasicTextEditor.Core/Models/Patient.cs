using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyBasicTextEditor.Core.Models
{
    public class Patient
    {
        private string forename;
        private string surname;
        private List<string> middleNames;
        private string fullName;
        private DateTime dob;
        private int patientNumber;
        private string initialledName;

        private string displayName;

        private Address primaryAddress;

        public Patient(string fname, string sname, DateTime dateOfBirth, int patNum, List<string> midNames = null)
        {
            this.Forename = fname;
            this.Surname = sname;
            this.DOB = dateOfBirth;
            this.PatientNumber = patNum;
            this.MiddleNames = midNames;




            this.FullName = GetFullName();
            this.DisplayName = string.Format("{0}, {1}", this.Surname.ToUpper(), this.Forename);
            this.InitialledName = string.Format("{0}. {1}{2}", this.Forename[0], this.MiddleNames != null ? MiddleInitials() : " ", this.Surname);

            this.PrimaryAddress = new Address(AddressType.Primary, "234", "Dundee", "Scotland", "DD2 12ZX", "Blackness Rd.", "Angus");
        }

        public string Forename
        {
            get { return forename; }
            set { forename = value; }
        }

        public string Surname
        {
            get { return surname; }
            set { surname = value; }
        }

        public List<string> MiddleNames
        {
            get { return middleNames; }
            set { middleNames = value; }
        }

        public string FullName
        {
            get { return fullName; }
            set { fullName = value; }
        }

        public DateTime DOB
        {
            get { return dob; }
            set { dob = value; }
        }

        public int PatientNumber
        {
            get { return patientNumber; }
            set { patientNumber = value; }
        }

        public string InitialledName
        {
            get { return initialledName; }
            set { initialledName = value; }
        }
        public string DisplayName
        {
            get { return displayName; }
            set { displayName = value; }
        }

        public Address PrimaryAddress
        {
            get { return primaryAddress; }
            set { primaryAddress = value; }
        }

        private string GetFullName()
        {
            List<string> fullNameString = new List<string>() { this.Forename, this.Surname };
            if (this.MiddleNames != null)
            {
                fullNameString.InsertRange(1, this.MiddleNames);
            }

            return string.Join(" ", fullNameString);
        }

        private string MiddleInitials()
        {
            string initials = string.Empty;

            foreach (string name in this.MiddleNames)
            {
                initials += name[0] + ". ";
            }

            return initials;
        }
    }
}
