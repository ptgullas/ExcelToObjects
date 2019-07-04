using System;
using System.Collections.Generic;
using System.Text;
using Serilog;

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
        public string State { get; set; }
        public string CellPhone { get; set; }
        public string HomePhone { get; set; }
        public string Email { get; set; }
        public DateTime DateOfBirth { get; set; }
        public string DateOfBirthStr {
            get {
                if (DateOfBirth > new DateTime(1900,1,1)) {
                    return DateOfBirth.ToString("d");
                }
                else {
                    return null;
                }
            }
        }
        // SHOULD NOT BE STORED AS A SEPARATE COLUMN, BUT WILL BE INCORPORATED IN AN EXISTING COLUMN:
        public string Apartment { get; set; }


        public void PadZipCodeWithZeroes() {
            // if the spreadsheet contained the Zip Code as a number, it may have removed
            // leading zeroes. This puts them back
            if (ZipCode != null) {
                if ((ZipCode.Length < 5) && (ZipCode.IsNumeric())) {
                    Log.Information("Padding Zip Code {zip}", ZipCode);
                    string fmt = "00000.##";
                    int ZipInt = ZipCode.ToInt();
                    ZipCode = ZipInt.ToString(fmt);
                }
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
            if (Address != null) {
                Address = Address.ReplaceInvalidChars();
            }
        }

        public void RemoveNonAlphanumericFromAddress() {
            if (Address != null) {
                Address = Address.RemoveNonAlphanumeric();
            }
        }

        public void ReplaceNumberSignInAddressWithApt() {
            if (Address != null) {
                Address = Address.Replace("#", "Apt ");
            }
        }

        public void RemoveMultipleSpacesFromAddress() {
            if (Address != null) {
                Address = Address.ReplaceWhitespaceWithSingleSpace();
            }
        }

        public void RemoveNonNumericAndSpacesFromPhones() {
            RemoveNonNumericFromCellPhone();
            RemoveNonNumericFromHomePhone();
        }

        public void SetHomePhoneToNullIfSameAsCellPhone() {
            if (!string.IsNullOrEmpty(CellPhone) && !string.IsNullOrEmpty(HomePhone)) {
                if (CellPhone == HomePhone) {
                    HomePhone = null;
                }
            }
        }

        public void ChangeZeroPhoneValuesToNull() {
            if (CellPhone == "0")
                CellPhone = null;
            if (HomePhone == "0")
                HomePhone = null;
        }

        private void RemoveNonNumericFromCellPhone() {
            if (CellPhone != null) {
                CellPhone = CellPhone.RemoveNonNumeric().RemoveWhitespace();
            }
        }

        private void RemoveNonNumericFromHomePhone() {
            if (HomePhone != null) {
                HomePhone = HomePhone.RemoveNonNumeric().RemoveWhitespace();
            }
        }

        public void AppendApartmentToAddress() {
            RemoveNAFromApartment();
            RemoveAptFromApartment();
            if (Apartment != null) {
                Address = $"{Address} Apt {Apartment}";
            }
        }

        private void RemoveNAFromApartment() {
            if (Apartment != null) {
                if (Apartment.RemoveNonAlphanumeric().ToLower() == "na") {
                    Apartment = null;
                }
            }
        }

        private void RemoveAptFromApartment() {
            if (Apartment != null) {
                Apartment = Apartment.Replace("Apt", "");
            }
        }

        public bool EmailIsValid() {
            return Email.IsValidEmail();
        }

    }



}
