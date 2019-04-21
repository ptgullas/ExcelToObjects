using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToObjects {
    public class Member {
        // REQUIRED:
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string ZipCode { get; set; }
        // OPTIONAL BUT VERY HELPFUL:
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public DateTime DateOfBirth { get; set; }
        public string MiddleName { get; set; }
        public string NameSuffix { get; set; }

        public bool ContainsFullAddress() {
            if ((!string.IsNullOrEmpty(Address))
                && (!string.IsNullOrEmpty(City))
                && (!string.IsNullOrEmpty(State))) {
                return true;
            }
            else {
                return false;
            }
        }

        public string GetFullAddress() {
            return $"{Address} {City} {State}";
        }



    }



}
