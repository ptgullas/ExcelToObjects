using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Linq;

namespace ExcelToObjects {
    class Program {
        static void Main(string[] args) {
            string myPath = @"C:\Users\Prime Time Pauly G\Documents\ProgHackNight TestAddresses.xlsx";
            Standardizer standardizer = new Standardizer(myPath);
        }



    }
}
