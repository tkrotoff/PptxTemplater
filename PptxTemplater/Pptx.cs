using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace PptxTemplater
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

        public Pptx(Stream stream, bool isEditable)
        {
            _pptx = PresentationDocument.Open(stream, isEditable);
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
            PresentationPart presentationPart = _pptx.PresentationPart;
            return presentationPart.SlideParts.Count();
        }

        /// Gets the PptxSlide given a slide index.
        private PptxSlide GetPptxSlide(int slideIndex)
        {
            PresentationPart presentationPart = _pptx.PresentationPart;

            // Get the collection of slide IDs
            OpenXmlElementList slideIds = presentationPart.Presentation.SlideIdList.ChildElements;

            // Get the relationship ID of the slide
            string relId = (slideIds[slideIndex] as SlideId).RelationshipId;

            // Get the specified slide part from the relationship ID
            SlidePart slide = (SlidePart) presentationPart.GetPartById(relId);

            return new PptxSlide(slide);
        }

        /// Gets all text found inside the given slide.
        ///
        /// See How to: Get All the Text in a Slide in a Presentation http://msdn.microsoft.com/en-us/library/office/cc850836
        /// See How to: Get All the Text in All Slides in a Presentation http://msdn.microsoft.com/en-us/library/office/gg278331
        public string[] GetAllTextInSlide(int slideIndex)
        {
            PptxSlide slide = GetPptxSlide(slideIndex);
            return slide.GetAllText();
        }

        /// Replaces a text (tag) by another inside the given slide.
        public void ReplaceTagInSlide(int slideIndex, string tag, string newText)
        {
            PptxSlide slide = GetPptxSlide(slideIndex);
            slide.ReplaceTag(tag, newText);
        }

        public void ReplacePictureInSlide(int slideIndex, string tag, string newPicture)
        {
            throw new System.NotImplementedException();
        }
    }
}
