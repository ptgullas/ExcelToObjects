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

        [Fact]
        public void ReplaceWhitespaceWithSingleSpace_ReplaceMultipleSpaces() {
            string strToTest = "2289    Broadway #3E";
            string expected = "2289 Broadway #3E";

            Assert.Equal(expected, strToTest.ReplaceWhitespaceWithSingleSpace());
        }
    }
}
