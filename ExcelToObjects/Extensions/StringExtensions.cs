using System;
using System.Collections.Generic;
using System.Globalization;
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
            // this is a colon, a backslash, and a forward slash, and a .
            string invalidChars = "[:\\\\/.]";
            return Regex.Replace(str, invalidChars, strToReplace);
        }

        public static string ReplaceWhitespaceWithSingleSpace(this string str) {
            return Regex.Replace(str, @"\s+", " ");
        }

        public static string RemoveWhitespace(this string str) {
            return str.Replace(" ", string.Empty);
        }

        public static string TrimIfNotNull(this string str) {
            if (!string.IsNullOrEmpty(str)) {
                return str.Trim();
            }
            else {
                return null;
            }
        }

        // taken from https://docs.microsoft.com/en-us/dotnet/standard/base-types/how-to-verify-that-strings-are-in-valid-email-format
        public static bool IsValidEmail(this string email) {
            if (string.IsNullOrWhiteSpace(email))
                return false;

            try {
                // Normalize the domain
                email = Regex.Replace(email, @"(@)(.+)$", DomainMapper,
                                      RegexOptions.None, TimeSpan.FromMilliseconds(200));

                // Examines the domain part of the email and normalizes it.
                string DomainMapper(Match match) {
                    // Use IdnMapping class to convert Unicode domain names.
                    var idn = new IdnMapping();

                    // Pull out and process domain name (throws ArgumentException on invalid)
                    var domainName = idn.GetAscii(match.Groups[2].Value);

                    return match.Groups[1].Value + domainName;
                }
            }
            catch (RegexMatchTimeoutException e) {
                return false;
            }
            catch (ArgumentException e) {
                return false;
            }

            try {
                return Regex.IsMatch(email,
                    @"^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                    @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-0-9a-z]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$",
                    RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250));
            }
            catch (RegexMatchTimeoutException) {
                return false;
            }
        }


    }
}
