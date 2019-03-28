﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using EPPlus.DataExtractor;

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
            if (worksheetNum > worksheet.Workbook.Worksheets.Count - 1) {
                return worksheet.GetHeaderColumns();
            }
            else {
                throw new ArgumentOutOfRangeException("Specified invalid worksheet number");
            }
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