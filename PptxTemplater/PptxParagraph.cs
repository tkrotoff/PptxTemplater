namespace PptxTemplater
{
    using System.Collections.Generic;
    using System.Text;
    using System.Text.RegularExpressions;

    using A = DocumentFormat.OpenXml.Drawing;

    /// <summary>
    /// Represents a paragraph inside a PowerPoint file.
    /// </summary>
    /// <remarks>Could not simply be named Paragraph, conflicts with DocumentFormat.OpenXml.Drawing.Paragraph.</remarks>
    internal static class PptxParagraph
    {
        internal static void ReplaceTag(A.Paragraph p, string tag, string newText)
        {
            while (true)
            {
                // Search for the tag
                Match match = Regex.Match(GetTexts(p), tag);
                if (!match.Success)
                {
                    break;
                }

                List<TextIndex> texts = GetTextIndexList(p);

                for (int i = 0; i < texts.Count; i++)
                {
                    TextIndex text = texts[i];
                    if (match.Index >= text.StartIndex && match.Index <= text.EndIndex)
                    {
                        // Got the right A.Text

                        int index = match.Index - text.StartIndex;
                        int done = 0;

                        for (; i < texts.Count; i++)
                        {
                            TextIndex currentText = texts[i];
                            List<char> currentTextChars = new List<char>(currentText.Text.Text.ToCharArray());

                            for (int k = index; k < currentTextChars.Count; k++, done++)
                            {
                                if (done < newText.Length)
                                {
                                    if (done >= tag.Length - 1)
                                    {
                                        // Case if newText is longer than the tag
                                        // Insert characters
                                        int remains = newText.Length - done;
                                        currentTextChars.RemoveAt(k);
                                        currentTextChars.InsertRange(k, newText.Substring(done, remains));
                                        done += remains;
                                        break;
                                    }
                                    else
                                    {
                                        currentTextChars[k] = newText[done];
                                    }
                                }
                                else
                                {
                                    if (done < tag.Length)
                                    {
                                        // Case if newText is shorter than the tag
                                        // Erase characters
                                        int remains = tag.Length - done;
                                        if (remains > currentTextChars.Count - k)
                                        {
                                            remains = currentTextChars.Count - k;
                                        }
                                        currentTextChars.RemoveRange(k, remains);
                                        done += remains;
                                        break;
                                    }
                                    else
                                    {
                                        // Regular case, nothing to do
                                        //currentTextChars[k] = currentTextChars[k];
                                    }
                                }
                            }

                            currentText.Text.Text = new string(currentTextChars.ToArray());
                            index = 0;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Returns all the texts found inside a given paragraph.
        /// </summary>
        /// <remarks>
        /// If all A.Text in the given paragraph are empty, returns an empty string.
        /// </remarks>
        internal static string GetTexts(A.Paragraph p)
        {
            StringBuilder concat = new StringBuilder();
            foreach (A.Text t in p.Descendants<A.Text>())
            {
                concat.Append(t.Text);
            }
            return concat.ToString();
        }

        /// <summary>
        /// Associates a A.Text with start and end index matching a paragraph full string (= the concatenation of all A.Text of a paragraph).
        /// </summary>
        private class TextIndex
        {
            public A.Text Text { get; private set; }
            public int StartIndex { get; private set; }
            public int EndIndex { get { return StartIndex + Text.Text.Length; } }

            public TextIndex(A.Text t, int startIndex)
            {
                this.Text = t;
                this.StartIndex = startIndex;
            }
        }

        /// <summary>
        /// Gets all the TextIndex for a given paragraph.
        /// </summary>
        private static List<TextIndex> GetTextIndexList(A.Paragraph p)
        {
            List<TextIndex> texts = new List<TextIndex>();

            StringBuilder concat = new StringBuilder();
            foreach (A.Text t in p.Descendants<A.Text>())
            {
                int startIndex = concat.Length;
                texts.Add(new TextIndex(t, startIndex));
                concat.Append(t.Text);
            }

            return texts;
        }
    }
}
