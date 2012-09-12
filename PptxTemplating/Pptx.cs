using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace PptxTemplating
{
    /// Represents a PowerPoint file.
    public class Pptx
    {
        private readonly PresentationDocument _pptx;

        /// Opens a PowerPoint file in read-write or read only mode.
        public Pptx(string file, bool isEditable)
        {
            _pptx = PresentationDocument.Open(file, isEditable);
        }

        /// Closes the PowerPoint file.
        /// 99% of the time this is not needed, the PowerPoint file will get closed when the destructor is being called.
        public void Close()
        {
            _pptx.Close();
        }

        /// Counts the number of slides in the presentation.
        ///
        /// See How to: Get All the Text in All Slides in a Presentation http://msdn.microsoft.com/en-us/library/office/gg278331
        public int CountSlides()
        {
            PresentationPart part = _pptx.PresentationPart;

            return part.SlideParts.Count();
        }

        /// Gets all text found inside the given slide.
        ///
        /// See How to: Get All the Text in a Slide in a Presentation http://msdn.microsoft.com/en-us/library/office/cc850836
        /// See How to: Get All the Text in All Slides in a Presentation http://msdn.microsoft.com/en-us/library/office/gg278331
        public string[] GetAllTextInSlide(int slideIndex)
        {
            PresentationPart part = _pptx.PresentationPart;

            // Get the collection of slide IDs
            OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

            // Get the relationship ID of the slide
            string relId = (slideIds[slideIndex] as SlideId).RelationshipId;

            // Get the specified slide part from the relationship ID
            SlidePart slide = (SlidePart) part.GetPartById(relId);

            return new PptxSlide(slide).GetAllText();
        }

        /// Replaces a text (tag) by another inside the given slide.
        public void ReplaceTagInSlide(int slideIndex, string tag, string newText)
        {
            PresentationPart part = _pptx.PresentationPart;

            // Get the collection of slide IDs
            OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

            // Get the relationship ID of the slide
            string relId = (slideIds[slideIndex] as SlideId).RelationshipId;

            // Get the specified slide part from the relationship ID
            SlidePart slide = (SlidePart) part.GetPartById(relId);

            new PptxSlide(slide).ReplaceTag(tag, newText);
        }
    }
}
