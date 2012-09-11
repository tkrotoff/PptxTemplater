using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
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
        /// See How to: Get All the Text in a Slide in a Presentation http://msdn.microsoft.com/en-us/library/office/cc850836
        public string[] GetAllText()
        {
            // Create a new linked list of strings.
            LinkedList<string> texts = new LinkedList<string>();

            // Iterate through all the paragraphs in the slide.
            foreach (A.Paragraph p in _slide.Slide.Descendants<A.Paragraph>())
            {
                StringBuilder paragraphText = new StringBuilder();

                // Iterate through the lines of the paragraph.
                foreach (A.Text t in p.Descendants<A.Text>())
                {
                    paragraphText.Append(t.Text);
                }

                if (paragraphText.Length > 0)
                {
                    texts.AddLast(paragraphText.ToString());
                }
            }

            return texts.ToArray();
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
                Hello this is a tag: <hello>
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
                Another tag: <bonjour
               </a:t>
              </a:r>
              <a:r>
               <a:rPr lang="en-US" dirty="0" smtClean="0"/>
               <a:t>
                > le monde !
               </a:t>
              </a:r>
              <a:endParaRPr lang="en-US" dirty="0"/>
             </a:p>
            */

            bool insert = true;
            //foreach (A.Paragraph p in _slide.Slide.Descendants<A.Paragraph>())
            for (int j = 0; j < _slide.Slide.Descendants<A.Paragraph>().Count(); j++)
            {
                A.Paragraph p = _slide.Slide.Descendants<A.Paragraph>().ElementAt(j);

                StringBuilder concat = new StringBuilder();
                List<int> splits = new List<int>();

                foreach (A.Run r in p.Descendants<A.Run>())
                {
                    foreach (A.Text t in r.Descendants<A.Text>())
                    {
                        string tmp = t.Text;
                        concat.Append(tmp);
                        splits.Add(tmp.Count());
                    }
                    if (insert)
                    {
                        //InsertTextInsideParagraph(p, r, "FUCK");
                        insert = false;
                    }
                }

                string fullText = concat.ToString();
                List<string> modifiedTexts = new List<string>();

                if (Regex.Match(fullText, tag).Success)
                {
                    string modifiedText = Regex.Replace(fullText, tag, newText);
                    modifiedTexts.AddRange(modifiedText.Substrings(splits));

                    var texts = p.Descendants<A.Text>().ToList();
                    for (int i = 0; i < texts.Count(); i++)
                    {
                        /*A.Text t = texts[i];
                        t.Text = modifiedTexts[i];*/
                        InsertTextInsideParagraph(p, i + 1, "YOU");
                    }
                }
            }
        }

        /// <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        ///   <a:r>
        ///     <a:rPr lang="en-US" dirty="0" smtClean="0" />
        ///     <a:t>Some text</a:t>
        ///   </a:r>
        ///   [...]
        /// </a:p>
        private static void InsertTextInsideParagraph(A.Paragraph p, A.Run rAfter, string text)
        {
            A.Run r = new A.Run();
            //A.RunProperties rPr = new A.RunProperties(/*rPrTemplate*/);
            A.RunProperties rPr = new A.RunProperties() { Language = "fr-FR", Dirty = false, SmartTagClean = false };
            A.Text t = new A.Text(text);

            r.AppendChild(rPr);
            //r.AppendChild(rPr);
            r.AppendChild(t);

            //p.AppendChild(r);
            p.InsertAfter(r, rAfter);
        }

        private static void InsertTextInsideParagraph(A.Paragraph p, int at, string text)
        {
            A.Run r = new A.Run();
            //A.RunProperties rPr = new A.RunProperties(/*rPrTemplate*/);
            A.RunProperties rPr = new A.RunProperties() { Language = "fr-FR", Dirty = false, SmartTagClean = false };
            A.Text t = new A.Text(text);

            r.AppendChild(rPr);
            //r.AppendChild(rPr);
            r.AppendChild(t);

            //p.AppendChild(r);
            p.InsertAt(r, at);
        }
    }
}
