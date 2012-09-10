using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace PptxTemplating
{
    public class Pptx
    {
        private readonly PresentationDocument _pptx;

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

            return new PptxSlide(slide).GetAllText();
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

            new PptxSlide(slide).ReplaceTag(tag, newText);
        }
    }
}
