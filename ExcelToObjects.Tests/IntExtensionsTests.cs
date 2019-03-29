using ExcelToObjects.Extensions;
using System;
using System.Collections.Generic;
using System.Text;
using Xunit;


namespace ExcelToObjects.Test {
    public class IntExtensionsTests {
        [Fact]
        public void ToLetter_NumberMatchesSingleLetter_ReturnsLetter() {
            string expectedLetter = "A";
            int numToTest = 1;

            Assert.Equal(expectedLetter, numToTest.ToLetter());
        }

        [Fact]
        public void ToLetter_NumberMatchesDoubleLetter_ReturnsLetter() {
            string expectedLetter = "AA";
            int numToTest = 27;

            Assert.Equal(expectedLetter, numToTest.ToLetter());
        }
        [Fact]
        public void ToLetter_NumberIsZero_ReturnsNull() {
            string expectedLetter = null;
            int numToTest = 0;

            Assert.Equal(expectedLetter, numToTest.ToLetter());
        }
    }
}
