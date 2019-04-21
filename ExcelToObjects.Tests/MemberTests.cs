using ExcelToObjects.Extensions;
using System;
using System.Collections.Generic;
using System.Text;
using Xunit;

namespace ExcelToObjects.Test {
    public class MemberTests {
        [Fact]
        public void ContainsFullAddress_DoesContain_ReturnsTrue() {
            Member member = new Member() {
                LastName = "Stark",
                FirstName = "Lyanna",
                Address = "2289 Broadway",
                City = "New York",
                State = "NY"
            };

            Assert.True(member.ContainsFullAddress());
        }

        [Fact]
        public void GetFullAddress_ReturnsAddress() {
            Member member = new Member() {
                LastName = "Stark",
                FirstName = "Lyanna",
                Address = "2289 Broadway",
                City = "New York",
                State = "NY"
            };
            string expected = "2289 Broadway New York NY";

            Assert.Equal(expected, member.GetFullAddress());
        }

        [Fact]
        public void PadZipCodeWithZeroes_HasLessThanFiveDigits_Pads() {
            Member member = new Member() {
                LastName = "Stark",
                FirstName = "Lyanna",
                Address = "2289 Broadway",
                City = "New York",
                State = "NY",
                ZipCode = "7345"
            };
            string expected = "07345";

            member.PadZipCodeWithZeroes();
            Assert.Equal(expected, member.ZipCode);
        }

        [Fact]
        public void PadZipCodeWithZeroes_HasFiveDigits_DoesNotPad() {
            Member member = new Member() {
                LastName = "Stark",
                FirstName = "Lyanna",
                Address = "2289 Broadway",
                City = "New York",
                State = "NY",
                ZipCode = "10024"
            };
            string expected = "10024";

            member.PadZipCodeWithZeroes();
            Assert.Equal(expected, member.ZipCode);
        }

        [Fact]
        public void ContainsFullAddressAndNoZip_ReturnsTrue() {
            Member member = new Member() {
                LastName = "Stark",
                FirstName = "Lyanna",
                Address = "2289 Broadway",
                City = "New York",
                State = "NY"
            };

            Assert.True(member.ContainsFullAddressAndNoZip());
        }

        [Fact]
        public void RemoveInvalidCharactersFromAddress_RemovesInvalid() {
            Member member = new Member() {
                LastName = "Stark",
                FirstName = "Lyanna",
                Address = "205 W. 95th St #23",
                City = "New York",
                State = "NY"
            };

            string expected = "205 W 95th St #23";

            member.RemoveInvalidCharactersFromAddress();
            Assert.Equal(expected, member.Address);
        }

        [Fact]
        public void ReplaceNumberSignInAddressWithApt_Replaces() {
            Member member = new Member() {
                LastName = "Stark",
                FirstName = "Lyanna",
                Address = "205 W. 95th St #23",
                City = "New York",
                State = "NY"
            };

            string expected = "205 W. 95th St Apt 23";

            member.ReplaceNumberSignInAddressWithApt();
            Assert.Equal(expected, member.Address);
        }

        [Fact]
        public void RemoveNonNumericFromPhones_Removes() {
            Member member = new Member() {
                LastName = "Stark",
                FirstName = "Lyanna",
                Address = "205 W. 95th St #23",
                City = "New York",
                State = "NY",
                CellPhone = "(917) 555-1212",
                HomePhone = "(718) 555-1223"
            };
            string expectedCell = "9175551212";

            member.RemoveNonNumericFromPhones();
            Assert.Equal(expectedCell, member.CellPhone);

        }


    }
}
