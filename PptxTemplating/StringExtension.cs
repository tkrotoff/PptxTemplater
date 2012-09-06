using System.Collections.Generic;

namespace PptxTemplating
{
    public static class StringExtension
    {
        public static string[] Substrings(this string str, List<int> lengths)
        {
            List<string> strList = new List<string>();

            int fullLength = 0;
            for (int i = 0; i < lengths.Count; i++)
            {
                int length = lengths[i];

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
