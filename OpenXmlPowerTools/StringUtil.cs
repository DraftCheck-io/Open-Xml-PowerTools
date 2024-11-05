using System;
using System.Collections.Generic;

namespace OpenXmlPowerTools
{
    public class StringUtil
    {

        private static readonly int[] CustomWhiteSpaceCharCodes = { 0xE905, 0xEFFE };

        public static bool IsWhiteSpaceChar(int charCode, int[] additionalWhiteSpaceCharCodes = null)
        {
            return charCode <= 0x20 ||
                charCode == -1 ||
                (additionalWhiteSpaceCharCodes != null && Array.Exists(additionalWhiteSpaceCharCodes, code => code == charCode)) ||
                Array.Exists(CustomWhiteSpaceCharCodes, code => code == charCode);
        }

        public static bool ContainsNonWhiteSpaceChar(string str, int[] customWhiteSpaceCharCodes = null)
        {
            if (string.IsNullOrEmpty(str)) return false;

            for (int i = 0; i < str.Length; i++)
            {
                if (!IsWhiteSpaceChar(str[i], customWhiteSpaceCharCodes))
                {
                    return true;
                }
            }
            return false;
        }

        public static string TrimStartWhiteSpaces(string str, int[] customWhiteSpaceCharCodes = null)
        {
            if (string.IsNullOrEmpty(str)) return str;

            int whiteSpaceCount = 0;
            for (int i = 0; i < str.Length; i++)
            {
                if (IsWhiteSpaceChar(str[i], customWhiteSpaceCharCodes))
                {
                    whiteSpaceCount++;
                }
                else
                {
                    break;
                }
            }
            return str.Substring(whiteSpaceCount);
        }

        public static string TrimEndWhiteSpaces(string str, int[] customWhiteSpaceCharCodes = null)
        {
            if (string.IsNullOrEmpty(str)) return str;

            int whiteSpaceCount = 0;
            for (int i = str.Length - 1; i >= 0; i--)
            {
                if (IsWhiteSpaceChar(str[i], customWhiteSpaceCharCodes))
                {
                    whiteSpaceCount++;
                }
                else
                {
                    break;
                }
            }
            str.TrimEnd();
            return str.Substring(0, str.Length - whiteSpaceCount);
        }

        public static string TrimWhiteSpaces(string str, int[] customWhiteSpaceCharCodes = null)
        {
            return TrimStartWhiteSpaces(TrimEndWhiteSpaces(str, customWhiteSpaceCharCodes), customWhiteSpaceCharCodes);
        }


        public class Word
        {
            public string Text { get; set; }
            public Range Range { get; set; }
        }

        public class Range
        {
            public int Start { get; set; }
            public int End { get; set; }
        }

        public static List<Word> GetWords(string text = "", bool upperCase = false)
        {
            var words = new List<Word>();

            int wordStartIndex = -1;

            // iterate till one char after the last one to add the last word
            for (int i = 0; i <= text.Length; i++)
            {
                int charCode = i < text.Length ? text[i] : -1;
                if (IsWhiteSpaceChar(charCode))
                {
                    if (wordStartIndex >= 0)
                    {
                        string word = text.Substring(wordStartIndex, i - wordStartIndex);
                        if (upperCase)
                        {
                            word = word.ToUpper();
                        }
                        words.Add(new Word
                        {
                            Text = word,
                            Range = new Range
                            {
                                Start = wordStartIndex,
                                End = i
                            }
                        });
                        wordStartIndex = -1;
                    }
                }
                else if (wordStartIndex < 0)
                {
                    wordStartIndex = i;
                }
            }

            return words;
        }

    }

}



