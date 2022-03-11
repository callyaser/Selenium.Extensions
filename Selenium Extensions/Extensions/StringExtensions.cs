using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace Selenium.Extensions
{
    public static class StringExtensions
    {
        public static string RemoveInvalidCharacters(this string str)
        {
            char[] permittedSpecialChars = new char[] { '_' };
            var stringArray = str.Where(c => char.IsLetterOrDigit(c) || permittedSpecialChars.Contains(c)).ToArray();
            return new string(stringArray);
        }
        public static bool IsNullOrEmpty(this string str)
        {
            return string.IsNullOrEmpty(str);
        }
        public static bool IsCyrillicOrChinese(this string str)
        {
            if (Regex.IsMatch(str, @"\p{IsCyrillic}"))
            {
                return true;
            }
            else if (str.Length > 1)
            {
                foreach (char c in str)
                {
                    UnicodeCategory uc = char.GetUnicodeCategory(c);
                    if (uc == UnicodeCategory.OtherLetter)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        public static bool In(this string targetEnv, params string[] strArray)
        {
            foreach (string str in strArray)
            {
                if (targetEnv.Equals(str))
                    return true;
            }
            return false;
        }
        public static bool NotIn(this string targetEnv, params string[] strArray)
        {
            foreach (string str in strArray)
            {
                if (targetEnv.Equals(str))
                    return false;
            }
            return true;
        }
        public static bool IsBetween(this int num, int min, int max)
        {
            return num >= min && num <= max;
        }
    }
}