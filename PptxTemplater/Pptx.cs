namespace PptxTemplater
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Drawing.Imaging;
    using System.IO;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Presentation;

    /// <summary>
    /// Represents a PowerPoint file.
    /// </summary>
    /// <returns>Follows the facade pattern.</returns>
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
        public int SlidesCount()
        {
            PresentationPart presentationPart = this.presentationDocument.PresentationPart;
            return presentationPart.SlideParts.Count();
        }

        /// <summary>
        /// Finds the slides matching a given note.
        /// </summary>
        /// <param name="note">Note to match the slide with.</param>
        /// <returns>The matching slides.</returns>
        public IEnumerable<PptxSlide> FindSlides(string note)
        {
            List<PptxSlide> slides = new List<PptxSlide>();

            for (int i = 0; i < this.SlidesCount(); i++)
            {
                PptxSlide slide = this.GetSlide(i);
                string[] notes = slide.GetNotes();
                if (notes.Contains(note))
                {
                    slides.Add(slide);
                }
            }

            return slides;
        }

        /// <summary>
        /// Gets the thumbnail (PNG format) associated with the PowerPoint file.
        /// </summary>
        /// <param name="size">The size of the thumbnail to generate, default is 256x192 pixels in 4:3 (160x256 in 16:10 portrait).</param>
        /// <returns>The thumbnail as a byte array (PNG format).</returns>
        /// <remarks>
        /// Even if the PowerPoint file does not contain any slide, still a thumbnail is generated.
        /// If the given size is superior to the default size then the thumbnail is upscaled and looks blurry so don't do it.
        /// </remarks>
        public byte[] GetThumbnail(Size size = default(Size))
        {
            byte[] thumbnail;

            var thumbnailPart = this.presentationDocument.ThumbnailPart;
            using (var stream = thumbnailPart.GetStream(FileMode.Open, FileAccess.Read))
            {
                var image = Image.FromStream(stream);
                if (size != default(Size))
                {
                    image = image.GetThumbnailImage(size.Width, size.Height, null, IntPtr.Zero);
                }

                using (var memoryStream = new MemoryStream())
                {
                    image.Save(memoryStream, ImageFormat.Png);
                    thumbnail = memoryStream.ToArray();
                }
            }

            return thumbnail;
        }

        /// <summary>
        /// Gets all the slides inside PowerPoint file.
        /// </summary>
        /// <returns>All the slides.</returns>
        public IEnumerable<PptxSlide> GetSlides()
        {
            List<PptxSlide> slides = new List<PptxSlide>();
            int nbSlides = this.SlidesCount();
            for (int i = 0; i < nbSlides; i++)
            {
                slides.Add(this.GetSlide(i));
            }
            return slides;
        }

        /// <summary>
        /// Gets the PptxSlide given a slide index.
        /// </summary>
        /// <param name="slideIndex">Index of the slide.</param>
        /// <returns>A PptxSlide.</returns>
        public PptxSlide GetSlide(int slideIndex)
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
