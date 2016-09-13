using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyBasicTextEditor.Core.Models
{
    public enum AddressType { Primary, Secondary, Holiday, Temporary };
    public class Address
    {
        private AddressType type;
        private string nameNum;
        private string streetName;
        private string town;
        private string county;
        private string country;
        private string postCode;

        private string dispalyAddress;

        public Address(AddressType addType, string addNumber, string addTown, string addCountry, string addPostCode, string addStName = null, string addCounty = null)
        {
            this.Type = addType;
            this.NameNum = addNumber;
            this.StreetName = addStName;
            this.Town = addTown;
            this.County = addCounty;
            this.Country = addCountry;
            this.PostCode = addPostCode;

            this.DisplayAddress = GetDisplayAddress();
        }

        public AddressType Type
        {
            get { return type; }
            set { type = value; }
        }

        public string NameNum
        {
            get { return nameNum; }
            set { nameNum = value; }
        }

        public string  StreetName
        {
            get { return streetName; }
            set { streetName = value; }
        }

        public string Town
        {
            get { return town; }
            set { town = value; }
        }

        public string County
        {
            get { return county; }
            set { county = value; }
        }

        public string Country
        {
            get { return country; }
            set { country = value; }
        }

        public string PostCode
        {
            get { return postCode; }
            set { postCode = value; }
        }

        public string DisplayAddress
        {
            get { return dispalyAddress; }
            set { dispalyAddress = value; }
        }

        private string GetDisplayAddress()
        {
            string newLine = " \n ";

            string addressFull = this.NameNum + " " + this.StreetName + newLine;
            addressFull += this.Town + newLine;
            addressFull += this.County != null ? this.County + newLine : null;
            addressFull += this.Country + newLine;
            addressFull += this.PostCode + newLine;

            return addressFull;
        }
    }
}
