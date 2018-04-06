using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpiraExcelAddIn
{
    public static class Utils
    {
        /// <summary>
        /// Returns a list of ids joined by a separator
        /// </summary>
        /// <param name="ids"></param>
        /// <param name="separator"></param>
        /// <returns></returns>
        public static string ToFormattedString(this int[] ids, string separator = ",")
        {
            if (ids == null || ids.Length < 1)
            {
                return "";
            }
            string str = "";
            foreach (int id in ids)
            {
                if (str == "")
                {
                    str = id.ToString();
                }
                else
                {
                    str += separator + id.ToString();
                }
            }
            return str;
        }
    }
}
