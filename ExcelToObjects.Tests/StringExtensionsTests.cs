using ExcelToObjects.Extensions;
using System;
using System.Collections.Generic;
using System.Text;
using Xunit;


namespace ExcelToObjects.Test {
    public class StringExtensionsTests {

        [Fact]
        public void IsValidEmail_ValidEmail_ReturnsTrue() {
            string emailToTest = "areacode212@gmail.com";
            bool result = emailToTest.IsValidEmail();

            Assert.True(result);
        }
    }
}
