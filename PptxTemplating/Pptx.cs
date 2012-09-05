using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
            foreach (A.Text text in texts)
            {
                paragraphText.Append(text.Text);
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
            foreach (DocumentFormat.OpenXml.Drawing.Paragraph paragraph in
                     slide.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
            {
                StringBuilder paragraphText = new StringBuilder();

                // Iterate through the lines of the paragraph.
                foreach (DocumentFormat.OpenXml.Drawing.Text text in
                         paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                {
                    paragraphText.Append(text.Text);
                }

                if (paragraphText.Length > 0)
                {
                    texts.AddLast(paragraphText.ToString());
                }
            }

            return texts.ToArray();
        }

    }
}
