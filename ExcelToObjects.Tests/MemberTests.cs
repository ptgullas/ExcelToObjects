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

    }
}
