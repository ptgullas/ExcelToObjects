using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Linq;
using Microsoft.Extensions.Configuration;
using Serilog;
using System.Threading.Tasks;

namespace ExcelToObjects {
    class Program {

        public static IConfigurationRoot Configuration;


        static async Task Main(string[] args) {

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Console()
                .WriteTo.File("c:\\temp\\logs\\ExcelToObjects\\ExcelToObjectsLog.txt", rollingInterval: RollingInterval.Day)
                .CreateLogger();
            Log.Information("Started ExcelToObjects log on {a}", DateTime.Now.ToLongTimeString());


            string projectRoot = AppContext.BaseDirectory.Substring(0, AppContext.BaseDirectory.LastIndexOf(@"\bin"));
            var builder = new ConfigurationBuilder()
                .SetBasePath(projectRoot)
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
            // IConfigurationRoot configuration = builder.Build();
            Configuration = builder.Build();

            try {
                //string GoogleApiKey = Configuration.GetValue<string>("GoogleApiKey");
                //ZipCodeRetrieverService zipRetriever = new ZipCodeRetrieverService(GoogleApiKey);
                //string myAddress = "225 e 17th street, new york ny";
                //string myZip = await zipRetriever.GetZip(myAddress);
                //Console.WriteLine($"Full Address is {myAddress} {myZip}");
                await ProcessSpreadsheets();
                Log.Information("Exiting nicely");
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
            }
        }

        private static async Task ProcessSpreadsheets() {
            string inputPath = Configuration.GetSection("Folders").GetValue<string>("inputFolder");
            string outputDir = @"C:\temp\ProgHackKnight_output";
            Directory.CreateDirectory(outputDir);
            if (Directory.Exists(inputPath)) {
                Log.Information("Processing folder {inputFolder}", inputPath);
                await ProcessFilesInInputFolder(inputPath, outputDir);
            }
            else {
                Log.Error<string>("Could not find folder {inputPath}", inputPath);
            }
        }

        private static async Task ProcessFilesInInputFolder(string inputFolder, string outputFolder)
        {
            List<string> inputSpreadsheets = GetFilePaths(inputFolder);
            Log.Information("Found {spreadsheetCount} spreadsheets", inputSpreadsheets.Count);
            foreach (string path in inputSpreadsheets) {
                try {
                    Log.Information("Processing spreadsheet {path}", path);
                    Standardizer standardizer = new Standardizer();
                    await ProcessSingleFile(path, outputFolder, standardizer);
                }
                catch (Exception e) {
                    Log.Error(e, "Error processing spreadsheet {spreadsheetName}", path);
                }
            }
        }

        private static List<string> GetFilePaths(string inputFolder)
        {
            return Directory.GetFiles(inputFolder, "*.xlsx").ToList();
        }

        private static async Task ProcessSingleFile(string myPath, string outputDir, Standardizer standardizer) {
            if (File.Exists(myPath)) {
                FileInfo sourceFile = new FileInfo(myPath);
                List<Member> members = new List<Member>();
                string newWorksheetName = null;
                using (ExcelPackage package = new ExcelPackage(sourceFile)) {
                    Log.Information("Retrieving members from {sourceFile}", sourceFile.Name);
                    members = standardizer.GetMembers(package, 0);
                    newWorksheetName = GetWorksheetName(package, 0);
                }

                members = await ProcessMembers(members);

                string newFilename = Path.GetFileNameWithoutExtension(myPath) + "_transformed.xlsx";
                string targetPath = Path.Combine(outputDir, newFilename);
                FileInfo targetFile = new FileInfo(targetPath);


                using (ExcelPackage targetPackage = new ExcelPackage(targetFile)) {
                    standardizer.ExportMembers(targetPackage, newWorksheetName, members);
                    targetPackage.SaveAs(targetFile);
                }
            }
        }

        static async Task<List<Member>> ProcessMembers(List<Member> members) {
            string GoogleApiKey = Configuration.GetValue<string>("GoogleApiKey");
            ZipCodeRetrieverService zipRetriever = new ZipCodeRetrieverService(GoogleApiKey);
            MemberProcessor memberProcessor = new MemberProcessor(zipRetriever);
            List<Member> newMembers = new List<Member>();

            // using Select like this lets us use i as the index
            foreach (Member m in members) {
                
                // Log.Information("Processing member {firstname} {lastname}", m.FirstName, m.LastName);
                m.TrimAllFields();
                m.PadZipCodeWithZeroes();
                m.ReplaceNumberSignInAddressWithApt();
                m.AppendApartmentToAddress();
                m.RemoveNonAlphanumericFromAddress();
                m.RemoveMultipleSpacesFromAddress();
                m.RemoveNonNumericAndSpacesFromPhones();
                m.ChangeZeroPhoneValuesToNull();
                m.SetHomePhoneToNullIfSameAsCellPhone();
                m.RemoveNAFromEmail();
                m.State = memberProcessor.GetStateAbbreviation(m);
                // m.ZipCode = await memberProcessor.GetZipFromMemberAddress(m);
                newMembers.Add(m);
                // Log.Information("Adding {firstName} {lastName} to newMembers", m.FirstName, m.LastName);
            }
            return newMembers;
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
