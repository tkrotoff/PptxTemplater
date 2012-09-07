using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxTemplating
{
    public class Pptx
    {
        private PresentationDocument _pptx;

        public Pptx(string file, bool isEditable)
        {
            _pptx = PresentationDocument.Open(file, isEditable);
        }

        public void Close()
        {
            _pptx.Close();
        }

        // Count the slides in the presentation.
        // See How to: Get All the Text in All Slides in a Presentation http://msdn.microsoft.com/en-us/library/office/gg278331
        public int CountSlides()
        {
            // Get the presentation part of document.
            PresentationPart part = _pptx.PresentationPart;

            return part.SlideParts.Count();
        }

        // See How to: Get All the Text in a Slide in a Presentation http://msdn.microsoft.com/en-us/library/office/cc850836
        // See How to: Get All the Text in All Slides in a Presentation http://msdn.microsoft.com/en-us/library/office/gg278331
        public string[] GetAllTextInSlide(int slideIndex)
        {
            // Get the presentation part of the presentation document.
            PresentationPart part = _pptx.PresentationPart;

            // Get the collection of slide IDs
            OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

            // Get the relationship ID of the slide.
            string relId = (slideIds[slideIndex] as SlideId).RelationshipId;

            // Get the specified slide part from the relationship ID.
            SlidePart slide = (SlidePart) part.GetPartById(relId);

            /*
            // Get the inner text of the slide:
            StringBuilder paragraphText = new StringBuilder();
            IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
            foreach (A.Text t in texts)
            {
                paragraphText.Append(t.Text);
            }
            */

            return GetAllTextInSlide(slide);
        }

        // See How to: Get All the Text in a Slide in a Presentation http://msdn.microsoft.com/en-us/library/office/cc850836
        private static string[] GetAllTextInSlide(SlidePart slide)
        {
            // Create a new linked list of strings.
            LinkedList<string> texts = new LinkedList<string>();

            // Iterate through all the paragraphs in the slide.
            foreach (A.Paragraph p in slide.Slide.Descendants<A.Paragraph>())
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

        public void ReplaceTagInSlide(int slideIndex, string tag, string newText)
        {
            // Get the presentation part of the presentation document.
            PresentationPart part = _pptx.PresentationPart;

            // Get the collection of slide IDs
            OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

            // Get the relationship ID of the slide.
            string relId = (slideIds[slideIndex] as SlideId).RelationshipId;

            // Get the specified slide part from the relationship ID.
            SlidePart slide = (SlidePart) part.GetPartById(relId);

            ReplaceTagInSlide(slide, tag, newText);
        }

        // See How to replace a paragraph's text using OpenXML SDK http://stackoverflow.com/questions/4276077/how-to-replace-an-paragraphs-text-using-openxml-sdk
        private void ReplaceTagInSlide(SlidePart slide, string tag, string newText)
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

            foreach (A.Paragraph p in slide.Slide.Descendants<A.Paragraph>())
            {
                StringBuilder concat = new StringBuilder();
                List<int> splits = new List<int>();

                foreach (A.Text t in p.Descendants<A.Text>())
                {
                    string tmp = t.Text;
                    concat.Append(tmp);
                    splits.Add(tmp.Count());
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
                        A.Text t = texts[i];
                        t.Text = modifiedTexts[i];
                    }
                }
            }
        }
    }
}
