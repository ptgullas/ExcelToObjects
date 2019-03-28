using System;
using System.IO;
using Xunit;
using OfficeOpenXml;
using System.Collections.Generic;

namespace ExcelToObjects.Tests {
    public class ExcelWorksheetExtensionsTests {
        [Fact]
        public void GetHeaderColumns_NormalHeaders_Passes() {
            string filePath = @"C:\Users\Prime Time Pauly G\Documents\ProgHackNight TestAddresses.xlsx";
            FileInfo spreadsheetFile = new FileInfo(filePath);
            string expectedfirstHeader = "Last";
            string expectedLastHeader = "E-Mail address";
            int expectedHeaderCount = 8;
            List<string> myHeaders;



            // Act
            using (ExcelPackage package = new ExcelPackage(spreadsheetFile)) {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                myHeaders = worksheet.GetHeaderColumns();
            }

            Assert.Equal(expectedHeaderCount, myHeaders.Count);
            Assert.Equal(expectedfirstHeader, myHeaders[0]);
            Assert.Equal(expectedLastHeader, myHeaders[myHeaders.Count - 1]);

        }
    }
}
