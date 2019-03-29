using System;
using System.IO;
using Xunit;
using OfficeOpenXml;
using System.Collections.Generic;

namespace ExcelToObjects.Tests {
    public class StandardizerTests {
        [Fact]
        public void GetColumnNumberOfFieldThatStartsWith_ColumnExists_ReturnsColNumber() {
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

        [Fact]
        public void GetHeaders_Worksheet0_ReturnsCorrectHeaders() {
            string filePath = @"C:\Users\Prime Time Pauly G\Documents\ProgHackNight TestAddresses.xlsx";
            FileInfo spreadsheetFile = new FileInfo(filePath);
            string expectedFirstColumn = "Last";
            string expectedLastColumn = "E-Mail address";

            Standardizer standardizer = new Standardizer(filePath);

            List<string> headers = new List<string>();

            using (ExcelPackage package = new ExcelPackage(spreadsheetFile)) {
                headers = standardizer.GetHeaders(package, 0);
            }

            Assert.Equal(expectedFirstColumn, headers[0]);
            Assert.Equal(expectedLastColumn, headers[7]);
        }



        [Fact]
        public void GetMembers_ValidMembers_ReturnsCorrectMembers() {
            string filePath = @"C:\Users\Prime Time Pauly G\Documents\ProgHackNight TestAddresses.xlsx";
            FileInfo spreadsheetFile = new FileInfo(filePath);
            string expectedFirst = "Aegon";
            string expectedLast = "Targaryen";
            string expectedZip = "10003";
            int expectedCount = 4;

            string expectedAddress = "51-38 Codwise Pl";

            Standardizer standardizer = new Standardizer(filePath);

            List<Member> members = new List<Member>();

            using (ExcelPackage package = new ExcelPackage(spreadsheetFile)) {
                members = standardizer.GetMembers(package, 0);
            }

            Assert.Equal(expectedFirst, members[0].FirstName);
            Assert.Equal(expectedLast, members[0].LastName);
            Assert.Equal(expectedZip, members[0].ZipCode);
            Assert.Equal(expectedAddress, members[1].Address);
            Assert.Equal(expectedCount, members.Count);
        }

        [Fact]
        public void GetMembers_MissingFields_ReturnsCorrectMembers() {
            string filePath = @"C:\Users\Prime Time Pauly G\Documents\ProgHackNight TestAddresses.xlsx";
            FileInfo spreadsheetFile = new FileInfo(filePath);
            string expectedFirst = "Tony";
            string expectedLast = "Stark";
            string expectedZip = null;
            int expectedCount = 4;

            Standardizer standardizer = new Standardizer(filePath);

            List<Member> members = new List<Member>();

            using (ExcelPackage package = new ExcelPackage(spreadsheetFile)) {
                members = standardizer.GetMembers(package, 1);
            }

            Assert.Equal(expectedFirst, members[2].FirstName);
            Assert.Equal(expectedLast, members[2].LastName);
            Assert.Equal(expectedZip, members[2].ZipCode);
            Assert.Equal(expectedCount, members.Count);
        }


        [Fact]
        public void GetColumnNumberOfFieldThatStartsWith_ColumnDoesNotExist_ReturnsZero() {
            string filePath = @"C:\Users\Prime Time Pauly G\Documents\ProgHackNight TestAddresses.xlsx";
            FileInfo spreadsheetFile = new FileInfo(filePath);
            Standardizer standardizer = new Standardizer(filePath);
            string fieldNameToSearch = "Favorite Weapon";
            int expectedColumnNumber = 0;

            List<string> myHeaders;

            // Act
            using (ExcelPackage package = new ExcelPackage(spreadsheetFile)) {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                myHeaders = worksheet.GetHeaderColumns();
            }

            int columnResult = standardizer.GetColumnNumberOfFieldThatStartsWith(myHeaders, fieldNameToSearch);
            Assert.Equal(expectedColumnNumber, columnResult);
        }

        [Fact]
        public void GetLastNameColumnNumber_ColumnExists_ReturnsColNumber() {
            string filePath = @"C:\Users\Prime Time Pauly G\Documents\ProgHackNight TestAddresses.xlsx";
            FileInfo spreadsheetFile = new FileInfo(filePath);
            Standardizer standardizer = new Standardizer(filePath);
            int expectedColumnNumber = 1;

            List<string> myHeaders;

            // Act
            using (ExcelPackage package = new ExcelPackage(spreadsheetFile)) {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                myHeaders = worksheet.GetHeaderColumns();
            }

            int columnResult = standardizer.GetLastNameColumnNumber(myHeaders);
            Assert.Equal(expectedColumnNumber, columnResult);

        }
    }
}
