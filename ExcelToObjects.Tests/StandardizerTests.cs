using System;
using System.IO;
using Xunit;
using OfficeOpenXml;
using System.Collections.Generic;

namespace ExcelToObjects.Tests {
    public class StandardizerTests {
        [Fact]
        public void GetColumnNumberOfFieldThatStartsWith_ColumnExists_Passes() {
            string filePath = @"C:\Users\Prime Time Pauly G\Documents\ProgHackNight TestAddresses.xlsx";
            FileInfo spreadsheetFile = new FileInfo(filePath);
            Standardizer standardizer = new Standardizer(filePath);
            string fieldNameToSearch = "Phone";
            int expectedColumnNumber = 7;

            List<string> myHeaders;

            // Act
            using (ExcelPackage package = new ExcelPackage(spreadsheetFile)) {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                myHeaders = worksheet.GetHeaderColumns();
            }

            int columnResult = standardizer.GetColumnNumberOfFieldThatStartsWith(myHeaders, fieldNameToSearch);
            Assert.Equal(expectedColumnNumber, columnResult);
        }
    }
}
