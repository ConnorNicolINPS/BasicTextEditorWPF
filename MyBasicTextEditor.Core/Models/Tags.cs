using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyBasicTextEditor.Core.Models
{
    public class Tags
    {
        public const string Forename="*PATFORENAME*";
        public const string Surname="*PATSURNAME*";
        public const string Middlenames="*PATMIDDLENAME*";
        public const string Fullname="*PATFULLNAME*";
        public const string Displayname="*PATDISPLAYNAME*";
        public const string Dateofbirth="*PATDATEORBIRTH*";
        public const string Idnumber="*PATNUMBER*";
        public const string MainAddress = "*PATMAINADDRESS*";
        public const string InitialledName = "*PATINITIALLEDNAME*";

        private string displayValue;
        private string tag;

        public Tags(string name, string tag)
        {
            this.DisplayValue = name;
            this.Tag = tag;
        }

        public string DisplayValue
        {
            get { return displayValue; }
            set { displayValue = value; }
        }

        public string Tag
        {
            get { return tag; }
            set { tag = value; }
        }



    }
}
