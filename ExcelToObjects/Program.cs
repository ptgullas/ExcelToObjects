using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Linq;

namespace ExcelToObjects {
    class Program {
        static void Main(string[] args) {
            try {
                string myPath = @"C:\Users\Prime Time Pauly G\Documents\ProgHackNight TestAddresses.xlsx";
                string outputDir = @"C:\temp\ProgHackKnight_output";
                Directory.CreateDirectory(outputDir);
                Standardizer standardizer = new Standardizer(myPath);
                if (File.Exists(myPath)) {
                    FileInfo sourceFile = new FileInfo(myPath);
                    List<Member> members = new List<Member>();
                    string newWorksheetName = null;
                    using (ExcelPackage package = new ExcelPackage(sourceFile)) {
                        members = standardizer.GetMembers(package, 0);
                        newWorksheetName = GetWorksheetName(package, 0);
                    }
                    string newFilename = Path.GetFileNameWithoutExtension(myPath) + "_transformed.xlsx";
                    string targetPath = Path.Combine(outputDir, newFilename);
                    FileInfo targetFile = new FileInfo(targetPath);

                    using (ExcelPackage targetPackage = new ExcelPackage(targetFile)) {
                        standardizer.ExportMembers(targetPackage, newWorksheetName, members);
                        targetPackage.SaveAs(targetFile);
                    }
                }
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
            }
        }

        static string GetWorksheetName(ExcelPackage package, int worksheetNum) {
            string name = null;
            using (package) {
                if (worksheetNum <= (package.Workbook.Worksheets.Count - 1)) {
                    name = package.Workbook.Worksheets[worksheetNum].Name;
                }
                else {
                    throw new ArgumentOutOfRangeException("worksheetNum", "Invalid workbook number");
                }
            }
            return name;
        }

    }
}
