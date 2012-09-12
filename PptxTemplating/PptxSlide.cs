using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxTemplating
{
    class PptxSlide
    {
        private readonly SlidePart _slide;

        public PptxSlide(SlidePart slide)
        {
            _slide = slide;
        }

        /// Returns all text found inside the slide.
        /// Some strings inside the array can be empty, this happens when all A.Text from a paragraph are empty
        /// See How to: Get All the Text in a Slide in a Presentation http://msdn.microsoft.com/en-us/library/office/cc850836
        public string[] GetAllText()
        {
            List<string> texts = new List<string>();
            foreach (A.Paragraph p in _slide.Slide.Descendants<A.Paragraph>())
            {
                texts.Add(GetParagraphAllText(p));
            }
            return texts.ToArray();
        }

        /// Returns all text found inside a given paragraph.
        /// If all A.Text in the given paragraph are empty, returns an empty string
        private string GetParagraphAllText(A.Paragraph p)
        {
            StringBuilder concat = new StringBuilder();
            foreach (A.Text t in p.Descendants<A.Text>())
            {
                concat.Append(t.Text);
            }
            return concat.ToString();
        }

        /// Associates a A.Text with start and end index matching a paragraph full string (= the concatenation of all A.Text of a paragraph).
        private class TextIndex
        {
            public A.Text Text { get; private set; }
            public int StartIndex { get; private set; }
            public int EndIndex { get { return StartIndex + Text.Text.Length; } }

            public TextIndex(A.Text t, int startIndex)
            {
                Text = t;
                StartIndex = startIndex;
            }
        }

        /// Gets all the TextIndex for a given paragraph.
        private List<TextIndex> GetTextIndexList(A.Paragraph p)
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

        /// Replaces a text (tag) by another inside the slide.
        /// See How to replace a paragraph's text using OpenXML SDK http://stackoverflow.com/questions/4276077/how-to-replace-an-paragraphs-text-using-openxml-sdk
        public void ReplaceTag(string tag, string newText)
        {
            /*
             <a:p>
              <a:r>
               <a:rPr lang="en-US" dirty="0" smtClean="0"/>
               <a:t>
                Hello this is a tag: {{hello}}
               </a:t>
              </a:r>
              <a:endParaRPr lang="fr-FR" dirty="0"/>
             </a:p>
            */

            /*
             <a:p>
              <a:r>
               <a:rPr lang="en-US" dirty="0" smtClean="0"/>
               <a:t>
                Another tag: {{bonjour
               </a:t>
              </a:r>
              <a:r>
               <a:rPr lang="en-US" dirty="0" smtClean="0"/>
               <a:t>
                }} le monde !
               </a:t>
              </a:r>
              <a:endParaRPr lang="en-US" dirty="0"/>
             </a:p>
            */

            foreach (A.Paragraph p in _slide.Slide.Descendants<A.Paragraph>())
            {
                while (true)
                {
                    string allText = GetParagraphAllText(p);

                    // Search for the tag
                    Match match = Regex.Match(allText, tag);
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
        }
    }
}
