using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Linq;
using Microsoft.Extensions.Configuration;

namespace ExcelToObjects {
    class Program {

        public static IConfigurationRoot Configuration;


        static void Main(string[] args) {
            string projectRoot = AppContext.BaseDirectory.Substring(0, AppContext.BaseDirectory.LastIndexOf(@"\bin"));
            var builder = new ConfigurationBuilder()
                .SetBasePath(projectRoot)
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
            // IConfigurationRoot configuration = builder.Build();
            Configuration = builder.Build();

            try {
                string inputPath = Configuration.GetSection("Folders").GetValue<string>("inputFolder");
                string myPath = @"C:\Users\Prime Time Pauly G\Documents\ProgHackNight TestAddresses.xlsx";
                string outputDir = @"C:\temp\ProgHackKnight_output";
                Directory.CreateDirectory(outputDir);
                ProcessFilesInInputFolder(inputPath, outputDir);
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
            }
        }

        private static void ProcessFilesInInputFolder(string inputFolder, string outputFolder)
        {
            List<string> inputSpreadsheets = GetFilePaths(inputFolder);
            foreach (string path in inputSpreadsheets)
            {
                Standardizer standardizer = new Standardizer();
                ProcessSingleFile(path, outputFolder, standardizer);
            }
        }

        private static List<string> GetFilePaths(string inputFolder)
        {
            return Directory.GetFiles(inputFolder, "*.xlsx").ToList();
        }

        private static void ProcessSingleFile(string myPath, string outputDir, Standardizer standardizer) {
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
