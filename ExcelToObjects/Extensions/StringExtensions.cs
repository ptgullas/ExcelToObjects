using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelToObjects {
    public static class StringExtensions {

        public static bool IsNumeric(this string source) {
            return int.TryParse(source, out int i);
        }

        public static int ToInt(this string source) {
            int result = 0;
            if (source.IsNumeric()) {
                bool b = int.TryParse(source, out result);
            }
            return result;
        }

        public static string RemoveNonNumeric(this string str) {
            char[] arr = str.Where(c => (char.IsDigit(c) ||
                             char.IsWhiteSpace(c)))
                            .ToArray();

            return new string(arr);
        }

        public static string RemoveNonAlphanumeric(this string str) {
            char[] arr = str.Where(c => (char.IsLetterOrDigit(c) ||
                             char.IsWhiteSpace(c)))
                            .ToArray();

            return new string(arr);
        }

        public static int GetLevenshteinDistance(this string str, string strToCompare) {
            return LevenshteinDistance.Compute(str, strToCompare);
        }

        public static string ReplaceInvalidChars(this string str, string strToReplace = "") {
            // this is a colon, a backslash, and a forward slash
            string invalidChars = "[:\\\\/]";
            return Regex.Replace(str, invalidChars, strToReplace);
        }

    }
}
