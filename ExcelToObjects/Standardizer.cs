using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using EPPlus.DataExtractor;
using ExcelToObjects.Extensions;

namespace ExcelToObjects {
    public class Standardizer {
        public string _filePath;
        private FileInfo _spreadsheetFile;
        public Standardizer(string filePath) {
            _filePath = filePath;
            _spreadsheetFile = new FileInfo(_filePath);
        }

        public List<string> GetHeaders(ExcelPackage package, int worksheetNum) {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetNum]; //worksheetNum starts at 0
            if (worksheetNum <= (package.Workbook.Worksheets.Count - 1)) {
                return worksheet.GetHeaderColumns();
            }
            else {
                throw new ArgumentOutOfRangeException("worksheetNum", "Invalid worksheetNum");
            }
        }

        public List<Member> GetMembers(ExcelPackage package, int worksheetNum = 0) {
            List<string> headers = GetHeaders(package, worksheetNum);
            ExcelWorksheet sheet = package.Workbook.Worksheets[worksheetNum];
            List<Member> members = sheet
                .Extract<Member>()
                .WithProperty(p => p.LastName, GetLastNameColumnNumber(headers).ToLetter())
                .WithProperty(p => p.FirstName, GetFirstNameColumnNumber(headers).ToLetter())
                .WithProperty(p => p.ZipCode, GetZipCodeColumnNumber(headers).ToLetter())
                .WithOptionalProperty(p => p.Address, GetAddressColumnNumber(headers).ToLetter())
                

                .GetData(2, sheet.Dimension.Rows)
                .ToList();
            return members;
        }

        // Last Name, First Name, Zip Code headers are all unlikely to start with anything else but those 3 words
        public int GetLastNameColumnNumber(List<string> headers) {
            return GetColumnNumberOfFieldThatStartsWith(headers, "Last");
        }

        public int GetFirstNameColumnNumber(List<string> headers) {
            return GetColumnNumberOfFieldThatStartsWith(headers, "First");
        }

        public int GetZipCodeColumnNumber(List<string> headers) {
            return GetColumnNumberOfFieldThatStartsWith(headers, "Zip");
        }

        public int GetAddressColumnNumber(List<string> headers) {
            int addressColumnNumber = GetColumnNumberOfFieldThatStartsWith(headers, "Add");
            if (addressColumnNumber == 0) {
                addressColumnNumber = GetColumnNumberOfFieldThatStartsWith(headers, "Street");
            }
            return addressColumnNumber;
        }

        public int GetColumnNumberOfFieldThatStartsWith(List<string> headers, string fieldNameToSearch) {
            var headersUpperCase = headers.Select(h => h.ToUpper()).ToList();
            string firstNameColText = headersUpperCase
                .FirstOrDefault(h => h.StartsWith(fieldNameToSearch.ToUpper()));
            if (firstNameColText != null) {
                int firstNameColNumber = headersUpperCase.IndexOf(firstNameColText);
                return firstNameColNumber + 1; // Excel columns (and rows) are 1-based, not zero-based
            }
            else {
                return 0;
            }

        }

        private static void PrintHeadersWithKnownNumberOfColumns(ExcelPackage package) {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int row = 1; // header row

            for (int column = 1; column <= 6; column++) {
                Console.WriteLine($"{worksheet.Cells[row, column].Text}");
            }
        }


    }
}
