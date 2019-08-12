using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Linq;
using Microsoft.Extensions.Configuration;
using Serilog;
using System.Threading.Tasks;
using ExcelToObjects.Extensions;

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

                if (args.Length == 0) {
                    DisplayHelpText();
                    await ProcessSpreadsheets();
                }
                else {
                    await ProcessArgs(args);
                }

                Log.Information("Exiting nicely");
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
            }
        }



        static async Task ProcessArgs(string[] args) {
            foreach (string arg in args) {
                switch(arg.Substring(0).ToUpper()) {
                    case "--HELP":
                        Console.WriteLine($"OK you correctly passed --help!");
                        DisplayHelpText();
                        break;
                    case "--PROCESS":
                        Console.WriteLine($"OK, you correctly passed --process!");
                        // await TestZipCodeRetriever();
                        await ProcessSpreadsheets();
                        break;
                    case "--COUNT":
                        Console.WriteLine($"OK, will start counting worksheets");
                        GetSpreadsheetsToCountWorksheets();
                        break;
                }
            }
        }
        static void DisplayHelpText() {
            string s = "ExcelToObjects.\n";
            s += "By Paul T. Gullas\n";
            s += "Display this help text: dotnet run -- --help\n";
            s += "dotnet run: Display help text then process spreadsheets\n";
            s += "Switches:\n";
            s += "dotnet run -- --count: Counts # of worksheets in spreadsheets in 'worksheetCountFolder'\n";
            s += "dotnet run -- --process: Process spreadsheets\n";
            Console.WriteLine(s);
        }

        private static void GetSpreadsheetsToCountWorksheets() {
            string path = Configuration.GetSection("Folders").GetValue<string>("worksheetCountFolder");
            if (Directory.Exists(path)) {
                List<string> spreadsheetsToCount = GetFilePaths(path);
                Log.Information("Found {spreadsheetCount} spreadsheets to count in {path}", spreadsheetsToCount.Count, path);
                foreach (string spreadsheet in spreadsheetsToCount) {
                    LogWorksheetCount(spreadsheet);
                }
            }
        }

        private static void LogWorksheetCount(string spreadsheet) {
            FileInfo sourceFile = new FileInfo(spreadsheet);
            using (ExcelPackage package = new ExcelPackage(sourceFile)) {
                int worksheetCount = package.Workbook.Worksheets.Count;
                string fileName = Path.GetFileName(spreadsheet);
                Log.Information("File {filename}: Worksheets: {worksheetCount}", fileName, worksheetCount);
            }
        }

        private static async Task TestZipCodeRetriever() {
            string GoogleApiKey = Configuration.GetValue<string>("GoogleApiKey");
            ZipCodeRetrieverService zipRetriever = new ZipCodeRetrieverService(GoogleApiKey);
            string myAddress = "225 e 17th street, new york ny";
            string myZip = await zipRetriever.GetZip(myAddress);
            Console.WriteLine($"Full Address is {myAddress} {myZip}");
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
                List<MembersInWorksheet> membersInSpreadsheet = GetMembersFromWorksheets(myPath, standardizer);

                await ProcessMembersInSpreadsheet(membersInSpreadsheet);

                ExportMembersToNewSpreadsheet(myPath, outputDir, standardizer, membersInSpreadsheet);
            }
        }

        private static void ExportMembersToNewSpreadsheet(string myPath, string outputDir, Standardizer standardizer, List<MembersInWorksheet> membersInSpreadsheet) {
            string newFilename = Path.GetFileNameWithoutExtension(myPath) + "_transformed.xlsx";
            string targetPath = Path.Combine(outputDir, newFilename);
            FileInfo targetFile = new FileInfo(targetPath);


            using (ExcelPackage targetPackage = new ExcelPackage(targetFile)) {
                foreach (MembersInWorksheet membersInWorksheet in membersInSpreadsheet) {
                    standardizer.ExportMembers(targetPackage, membersInWorksheet.NewWorksheetName, membersInWorksheet.Members);
                }
                targetPackage.SaveAs(targetFile);
            }
        }

        private static List<MembersInWorksheet> GetMembersFromWorksheets(string myPath, Standardizer standardizer) {
            FileInfo sourceFile = new FileInfo(myPath);
            List<MembersInWorksheet> membersInSpreadsheet = new List<MembersInWorksheet>();
            using (ExcelPackage package = new ExcelPackage(sourceFile)) {
                Log.Information("Retrieving members from {sourceFile}", sourceFile.Name);
                int worksheetCount = package.Workbook.Worksheets.Count;
                for (int currentWorksheet = 0; currentWorksheet < worksheetCount; currentWorksheet++) {
                    MembersInWorksheet membersInWorksheet = new MembersInWorksheet();
                    string worksheetName = GetWorksheetName(package.Workbook, currentWorksheet);
                    Log.Information("Retrieving members from worksheet {worksheetName}", worksheetName);
                    membersInWorksheet.Members = standardizer.GetMembers(package, currentWorksheet);
                    membersInWorksheet.NewWorksheetName = worksheetName;
                    membersInSpreadsheet.Add(membersInWorksheet);
                }
            }

            return membersInSpreadsheet;
        }

        private static async Task ProcessMembersInSpreadsheet(List<MembersInWorksheet> membersInSpreadsheet) {
            foreach (MembersInWorksheet membersInWorksheet in membersInSpreadsheet) {
                membersInWorksheet.Members = await ProcessMembers(membersInWorksheet.Members);
            }
        }

        static async Task<List<Member>> ProcessMembers(List<Member> members) {
            string GoogleApiKey = Configuration.GetValue<string>("GoogleApiKey");
            ZipCodeRetrieverService zipRetriever = new ZipCodeRetrieverService(GoogleApiKey);
            MemberProcessor memberProcessor = new MemberProcessor(zipRetriever);
            List<Member> newMembers = new List<Member>();

            // this lets us capture the index
            // borrowed from here: https://stackoverflow.com/a/39997157/11199987
            foreach (var (m, index) in members.WithIndex()) {
                
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
                m.State = memberProcessor.GetStateAbbreviation(m, index);
                // m.ZipCode = await memberProcessor.GetZipFromMemberAddress(m);
                newMembers.Add(m);
                // Log.Information("Adding {firstName} {lastName} to newMembers", m.FirstName, m.LastName);
            }
            return newMembers;
        }

        static string GetWorksheetName(ExcelWorkbook workbook, int worksheetNum) {
            string name = null;
            if (worksheetNum <= (workbook.Worksheets.Count - 1)) {
                name = workbook.Worksheets[worksheetNum].Name;
            }
            else {
                throw new ArgumentOutOfRangeException("worksheetNum", "Invalid workbook number");
            }
            return name;
        }

    }
}
