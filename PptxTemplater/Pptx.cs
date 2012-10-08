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
        /// Gets all the texts found inside the given slide.
        /// </summary>
        /// <param name="slideIndex">Index of the slide.</param>
        /// <returns>The texts inside a specific slide.</returns>
        /// <see href="http://msdn.microsoft.com/en-us/library/office/cc850836">How to: Get All the Text in a Slide in a Presentation</see>
        /// <see href="http://msdn.microsoft.com/en-us/library/office/gg278331">How to: Get All the Text in All Slides in a Presentation</see>
        /// <remarks>Internal method: needed for the unit tests only.</remarks>
        public string[] GetTextsInSlide(int slideIndex)
        {
            PptxSlide slide = this.GetPptxSlide(slideIndex);
            return slide.GetTexts();
        }

        /// <summary>
        /// Gets all the notes found inside the given slide.
        /// </summary>
        /// <param name="slideIndex">Index of the slide.</param>
        /// <returns>The notes inside a specific slide.</returns>
        public string[] GetNotesInSlide(int slideIndex)
        {
            PptxSlide slide = this.GetPptxSlide(slideIndex);
            return slide.GetNotes();
        }

        /// <summary>
        /// Gets all the tables found inside the given slide.
        /// </summary>
        /// <param name="slideIndex">Index of the slide.</param>
        /// <returns>The tables inside a specific slide.</returns>
        public PptxTable[] GetTablesInSlide(int slideIndex)
        {
            PptxSlide slide = this.GetPptxSlide(slideIndex);
            return slide.GetTables();
        }

        /// <summary>
        /// Finds all the tables that match the given tag.
        /// </summary>
        public PptxTable[] FindTables(string tag)
        {
            List<PptxTable> tables = new List<PptxTable>();

            for (int i = 0; i < this.SlidesCount(); i++)
            {
                PptxSlide slide = this.GetPptxSlide(i);
                tables.AddRange(slide.FindTables(tag));
            }

            return tables.ToArray();
        }

        /// <summary>
        /// Replaces a text (tag) by another inside the given slide.
        /// </summary>
        /// <remarks>Always call this method before PptxTable.SetRows() otherwise the number of slides might change.</remarks>
        /// <param name="slideIndex">Index of the slide.</param>
        /// <param name="tag">The tag to replace by newText, if null or empty do nothing; tag is a regex string.</param>
        /// <param name="newText">The new text to replace the tag with, if null replaced by empty string.</param>
        public void ReplaceTagInSlide(int slideIndex, string tag, string newText)
        {
            PptxSlide slide = this.GetPptxSlide(slideIndex);
            slide.ReplaceTag(tag, newText);
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
        /// Replaces a picture by another from a file inside the given slide.
        /// </summary>
        /// <remarks>Always call this method before PptxTable.SetRows() otherwise the number of slides might change.</remarks>
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
        /// <remarks>Always call this method before PptxTable.SetRows() otherwise the number of slides might change.</remarks>
        /// <param name="slideIndex">Index of the slide.</param>
        /// <param name="tag">The tag.</param>
        /// <param name="newPicture">The new picture.</param>
        /// <param name="contentType">Type of the content (image/png, image/jpeg...).</param>
        public void ReplacePictureInSlide(int slideIndex, string tag, Stream newPicture, string contentType)
        {
            PptxSlide slide = this.GetPptxSlide(slideIndex);
            slide.ReplacePicture(tag, newPicture, contentType);
        }

        /// <summary>
        /// Removes the given slide from the final PowerPoint file.
        /// </summary>
        /// <param name="slideIndex">Index of the slide.</param>
        public void RemoveSlide(int slideIndex)
        {
            PptxSlide slide = this.GetPptxSlide(slideIndex);
            slide.Remove();
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
