using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using EPPlus.DataExtractor;
using ExcelToObjects.Extensions;
using System.Drawing;

namespace ExcelToObjects {
    public class Standardizer {
        //public string _filePath;
        //private FileInfo _spreadsheetFile;
        public Standardizer() {
            //_filePath = filePath;
            //_spreadsheetFile = new FileInfo(_filePath);
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
                // ZipCode is technically required (and would normally use WithProperty), 
                // but if it's missing in the spreadsheet, 
                // then we will try & populate it using Address, City & State
                // But, on the other hand, I think we might expect ZipCode to actually 
                // be a column header; it just may not be populated on every field.
                .WithOptionalProperty(p => p.ZipCode, GetZipCodeColumnNumber(headers).ToLetter())
                .WithOptionalProperty(p => p.Address, GetAddressColumnNumber(headers).ToLetter())
                .WithOptionalProperty(p => p.City, GetCityColumnNumber(headers).ToLetter())
                .WithOptionalProperty(p => p.State, GetStateColumnNumber(headers).ToLetter())

                .GetData(2, sheet.Dimension.Rows)
                .ToList();
            return members;
        }

        public void ExportMembers(ExcelPackage package, string worksheetName, List<Member> members) {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetName);
            PopulateHeaders(worksheet);
            FormatHeaders(worksheet);
            PopulateMembers(members, worksheet);
        }

        private static void PopulateHeaders(ExcelWorksheet worksheet) {
            worksheet.Cells[1, 1].Value = "Last Name";
            worksheet.Cells[1, 2].Value = "First Name";
            worksheet.Cells[1, 3].Value = "Middle Name";
            worksheet.Cells[1, 4].Value = "Suffix";
            worksheet.Cells[1, 5].Value = "Street Address";
            worksheet.Cells[1, 6].Value = "City";
            worksheet.Cells[1, 7].Value = "State";
            worksheet.Cells[1, 8].Value = "Zip Code";
            worksheet.Cells[1, 9].Value = "Phone";
            worksheet.Cells[1, 10].Value = "E-mail";
            worksheet.Cells[1, 11].Value = "Date of Birth";
        }

        private static void FormatHeaders(ExcelWorksheet worksheet) {
            using (ExcelRange range = worksheet.Cells["A1:K1"]) {
                range.Style.Font.Bold = true;
                range.Style.Font.Color.SetColor(Color.FromArgb(217, 225, 242));
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(32, 55, 100));
            }
        }

        private static void PopulateMembers(List<Member> members, ExcelWorksheet worksheet) {
            if (members.Count > 0) {
                int rowStart = 2;
                int row = rowStart;
                foreach (Member m in members) {
                    worksheet.Cells[row, 1].Value = m.LastName;
                    worksheet.Cells[row, 2].Value = m.FirstName;
                    worksheet.Cells[row, 3].Value = m.MiddleName;
                    worksheet.Cells[row, 4].Value = m.NameSuffix;
                    worksheet.Cells[row, 5].Value = m.Address;
                    worksheet.Cells[row, 6].Value = m.City;
                    worksheet.Cells[row, 7].Value = m.State;
                    worksheet.Cells[row, 8].Value = m.ZipCode;
                    worksheet.Cells[row, 9].Value = m.Phone;
                    worksheet.Cells[row, 10].Value = m.Email;
                    worksheet.Cells[row, 11].Value = m.DateOfBirth;
                    row++;
                }
            }
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

        public int GetCityColumnNumber(List<string> headers) {
            return GetColumnNumberOfFieldThatStartsWith(headers, "City");
        }

        public int GetStateColumnNumber(List<string> headers) {
            return GetColumnNumberOfFieldThatStartsWith(headers, "St");
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
