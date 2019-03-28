using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;

namespace ExcelToObjects {
    public static class ExcelWorksheetExtensions {
        public static List<string> GetHeaderColumns(this ExcelWorksheet sheet) {
            List<string> columnNames = new List<string>();
            foreach (var firstRowCell in sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, 1, sheet.Dimension.End.Column]) {
                columnNames.Add(firstRowCell.Text);
            }
            return columnNames;
        }
    }
}
