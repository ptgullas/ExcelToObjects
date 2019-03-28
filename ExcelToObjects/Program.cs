using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace ExcelToObjects {
    class Program {
        static void Main(string[] args) {
            string myPath = @"C:\Users\Prime Time Pauly G\Documents\ProgHackNight TestAddresses.xlsx";
            ReadSpreadsheet(myPath);
        }

        static void ReadSpreadsheet(string filePath) {
            Console.WriteLine($"Reading file {filePath}");
            FileInfo spreadsheetFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(spreadsheetFile)) {
                PrintHeadersWithKnownNumberOfColumns(package);
            }
        }

        private static void PrintHeadersWithKnownNumberOfColumns(ExcelPackage package) {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int row = 1; // header row

            for (int column = 1; column <= 6; column++) {
                Console.WriteLine($"{worksheet.Cells[row, column].Text}");
            }
        }

        static void ReadHeaders(ExcelPackage package) {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            List<string> headers = worksheet.GetHeaderColumns();
        }

    }
}
