using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToObjects.Extensions {
    public static class IntExtensions {
        /// <summary>
        /// Convert an int to an Excel column name string (1 = A, 2 = B,..., 27 = AA, 28 = AB, etc)
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public static string ToLetter(this int i) {
            string result = string.Empty;
            while (--i >= 0) {
                result = (char)('A' + i % 26) + result;
                i /= 26;
            }
            return result;
        }
    }
}
