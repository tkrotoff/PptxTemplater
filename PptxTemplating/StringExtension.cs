using System.Collections.Generic;

namespace PptxTemplating
{
    public static class StringExtension
    {
        // Splits a string into several substrings given a list of substring lengths.
        // Example:
        // string str = "Bonjour tout le monde";
        // int[] splits = { 5, 8, 3 };
        // string[] substrings = str.Substrings(splits);
        // Assert.AreEqual(3, substrings.Length);
        // Assert.AreEqual("Bonjo", substrings[0]);
        // Assert.AreEqual("ur tout ", substrings[1]);
        // Assert.AreEqual("le ", substrings[2]);
        public static string[] Substrings(this string str, IEnumerable<int> lengths)
        {
            List<string> strList = new List<string>();

            int fullLength = 0;
            foreach (int length in lengths)
            {
                if (fullLength + length >= str.Length)
                {
                    strList.Add(str.Substring(fullLength, str.Length - fullLength));
                    fullLength = str.Length;
                }
                else
                {
                    strList.Add(str.Substring(fullLength, length));
                    fullLength += length;
                }
            }

            return strList.ToArray();
        }
    }
}
