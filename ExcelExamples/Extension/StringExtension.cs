using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelExamples.Extension {

    public static class StringExtensions {

        /// <summary>
        /// returns the index of the first digit.
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static int FirstDigitIndex(this String str) {
            for (int i = 0; i < str.Length; i++) {
                if (Char.IsDigit(str[i]))
                    return i;
            }
            return -1;
        }

        /// <summary>
        /// returns the numeric portion of a string
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static uint GetNumericValue(this String str) {
            return Convert.ToUInt32(Regex.Match(str, @"\d+").Value);
        }
    }



}
