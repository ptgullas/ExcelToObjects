using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToObjects {
    public class Member {
        // REQUIRED:
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string ZipCode { get; set; }
        // OPTIONAL BUT VERY HELPFUL:
        public string MiddleName { get; set; }
        public string NameSuffix { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        private string state;
        public string State { get { return state;  } set { state = value.ToUpper(); } }
        public string CellPhone { get; set; }
        public string HomePhone { get; set; }
        public string Email { get; set; }
        public DateTime DateOfBirth { get; set; }

        public void PadZipCodeWithZeroes() {
            // if the spreadsheet contained the Zip Code as a number, it may have removed
            // leading zeroes. This puts them back
            if ((ZipCode.Length < 5) && (ZipCode.IsNumeric())) {
                string fmt = "00000.##";
                int ZipInt = ZipCode.ToInt();
                ZipCode = ZipInt.ToString(fmt);
            }
        }

        public bool ContainsFullAddressAndNoZip() {
            bool result = false;
            if (ContainsFullAddress() && string.IsNullOrEmpty(ZipCode)) {
                result = true;
            }
            return result;
        }

        public bool ContainsFullAddress() {
            bool result = false;
            if ((!string.IsNullOrEmpty(Address)) && (!string.IsNullOrEmpty(City)) && (!string.IsNullOrEmpty(State))) {
                result = true;
            }
            return result;
        }

        public string GetFullAddress() {
            return $"{Address} {City} {State}";
        }

        public void RemoveInvalidCharactersFromAddress() {
            Address = Address.ReplaceInvalidChars();
        }

        public void RemoveNonAlphanumericFromAddress() {
            Address = Address.RemoveNonAlphanumeric();
        }

        public void ReplaceNumberSignInAddressWithApt() {
            Address = Address.Replace("#", "Apt ");
        }

        public void RemoveMultipleSpacesFromAddress() {
            Address = Address.ReplaceWhitespaceWithSingleSpace();
        }

        public void RemoveNonNumericFromPhones() {
            CellPhone = CellPhone.RemoveNonNumeric().RemoveWhitespace();
            HomePhone = HomePhone.RemoveNonNumeric().RemoveWhitespace();
        }

        public bool EmailIsValid() {
            return Email.IsValidEmail();
        }

    }



}
