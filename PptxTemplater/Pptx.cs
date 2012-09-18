namespace PptxTemplater
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Presentation;

    /// <summary>
    /// Represents a PowerPoint file.
    /// </summary>
    public class Pptx
    {
        private readonly PresentationDocument presentationDocument;

        #region ctor

        /// <summary>
        /// Initializes a new instance of the <see cref="Pptx"/> class.
        /// </summary>
        /// <param name="file">The PowerPoint file.</param>
        /// <param name="isEditable"><c>true</c> for read-write mode, <c>false</c> for read only mode.</param>
        /// <remarks>Opens a PowerPoint file in read-write (default) or read only mode.</remarks>
        public Pptx(string file, bool isEditable = true)
        {
            this.presentationDocument = PresentationDocument.Open(file, isEditable);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Pptx"/> class.
        /// </summary>
        /// <param name="stream">The PowerPoint stream.</param>
        /// <param name="isEditable"><c>true</c> for read-write mode, <c>false</c> for read only mode.</param>
        /// <remarks>Opens a PowerPoint stream in read-write (default) or read only mode.</remarks>
        public Pptx(Stream stream, bool isEditable = true)
        {
            this.presentationDocument = PresentationDocument.Open(stream, isEditable);
        }

        #endregion ctor

        /// <summary>
        /// Closes the PowerPoint file.
        /// </summary>
        /// <remarks>
        /// 99% of the time this is not needed, the PowerPoint file will get closed when the destructor is being called.
        /// </remarks>
        public void Close()
        {
            this.presentationDocument.Close();
        }

        /// <summary>
        /// Counts the number of slides.
        /// </summary>
        /// <returns>The number of slides.</returns>
        /// <see href="http://msdn.microsoft.com/en-us/library/office/gg278331">How to: Get All the Text in All Slides in a Presentation</see>
        public int CountSlides()
        {
            PresentationPart presentationPart = this.presentationDocument.PresentationPart;
            return presentationPart.SlideParts.Count();
        }

        /// <summary>
        /// Gets all text found inside the given slide.
        /// </summary>
        /// <param name="slideIndex">Index of the slide.</param>
        /// <returns>The text inside a specific slide.</returns>
        /// <see href="http://msdn.microsoft.com/en-us/library/office/cc850836">How to: Get All the Text in a Slide in a Presentation</see>
        /// <see href="http://msdn.microsoft.com/en-us/library/office/gg278331">How to: Get All the Text in All Slides in a Presentation</see>
        public string[] GetAllTextInSlide(int slideIndex)
        {
            PptxSlide slide = this.GetPptxSlide(slideIndex);
            return slide.GetAllText();
        }

        /// <summary>
        /// Replaces a text (tag) by another inside the given slide.
        /// </summary>
        /// <param name="slideIndex">Index of the slide.</param>
        /// <param name="tag">The tag.</param>
        /// <param name="newText">The new text.</param>
        public void ReplaceTagInSlide(int slideIndex, string tag, string newText)
        {
            PptxSlide slide = this.GetPptxSlide(slideIndex);
            slide.ReplaceTag(tag, newText);
        }

        /// <summary>
        /// Replaces a picture by another from a file inside the given slide.
        /// </summary>
        /// <param name="slideIndex">Index of the slide.</param>
        /// <param name="tag">The tag.</param>
        /// <param name="newPictureFile">The new picture file.</param>
        /// <param name="contentType">Type of the content (image/png, image/jpeg...).</param>
        public void ReplacePictureInSlide(int slideIndex, string tag, string newPictureFile, string contentType)
        {
            using (FileStream stream = new FileStream(newPictureFile, FileMode.Open, FileAccess.Read))
            {
                this.ReplacePictureInSlide(slideIndex, tag, stream, contentType);
            }
        }

        /// <summary>
        /// Replaces a picture by another from a stream inside the given slide.
        /// </summary>
        /// <param name="slideIndex">Index of the slide.</param>
        /// <param name="tag">The tag.</param>
        /// <param name="newPicture">The new picture.</param>
        /// <param name="contentType">Type of the content (image/png, image/jpeg...).</param>
        public void ReplacePictureInSlide(int slideIndex, string tag, Stream newPicture, string contentType)
        {
            PptxSlide slide = this.GetPptxSlide(slideIndex);
            slide.ReplacePicture(tag, newPicture, contentType);
        }

        public PptxTable[] FindTables(string tag)
        {
            List<PptxTable> tables = new List<PptxTable>();

            for (int i = 0; i < this.CountSlides(); i++)
            {
                PptxSlide slide = this.GetPptxSlide(i);
                tables.AddRange(slide.FindTables(tag));
            }

            return tables.ToArray();
        }

        /// <summary>
        /// Gets the PptxSlide given a slide index.
        /// </summary>
        /// <param name="slideIndex">Index of the slide.</param>
        /// <returns>A PptxSlide</returns>
        private PptxSlide GetPptxSlide(int slideIndex)
        {
            PresentationPart presentationPart = this.presentationDocument.PresentationPart;

            // Get the collection of slide IDs
            OpenXmlElementList slideIds = presentationPart.Presentation.SlideIdList.ChildElements;

            // Get the relationship ID of the slide
            string relId = ((SlideId)slideIds[slideIndex]).RelationshipId;

            // Get the specified slide part from the relationship ID
            SlidePart slidePart = (SlidePart)presentationPart.GetPartById(relId);

            return new PptxSlide(presentationPart, slidePart);
        }
    }
}
