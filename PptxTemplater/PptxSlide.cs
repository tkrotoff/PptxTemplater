namespace PptxTemplater
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Presentation;

    using A = DocumentFormat.OpenXml.Drawing;
    using Picture = DocumentFormat.OpenXml.Presentation.Picture;

    /// <summary>
    /// Represents a slide inside a PowerPoint file.
    /// </summary>
    /// <remarks>Could not simply be named Slide, conflicts with DocumentFormat.OpenXml.Drawing.Slide.</remarks>
    public class PptxSlide
    {
        private readonly PresentationPart presentationPart;
        private readonly SlidePart slidePart;

        internal PptxSlide(PresentationPart presentationPart, SlidePart slidePart)
        {
            this.presentationPart = presentationPart;
            this.slidePart = slidePart;
        }

        /// <summary>
        /// Gets all the texts found inside the slide.
        /// </summary>
        /// <remarks>
        /// Some strings inside the array can be empty, this happens when all A.Text from a paragraph are empty
        /// <see href="http://msdn.microsoft.com/en-us/library/office/cc850836">How to: Get All the Text in a Slide in a Presentation</see>
        /// </remarks>
        public IEnumerable<string> GetTexts()
        {
            List<string> texts = new List<string>();
            foreach (A.Paragraph p in this.slidePart.Slide.Descendants<A.Paragraph>())
            {
                texts.Add(PptxParagraph.GetTexts(p));
            }
            return texts;
        }

        /// <summary>
        /// Gets the slide title if any.
        /// </summary>
        /// <returns>The title or an empty string.</returns>
        public string GetTitle()
        {
            string title = string.Empty;

            // Find the title if any
            Shape titleShape = this.slidePart.Slide.Descendants<Shape>().FirstOrDefault(shape => IsShapeATitle(shape));
            if (titleShape != null)
            {
                title = string.Join(" ", titleShape.Descendants<A.Paragraph>().Select(paragraph => PptxParagraph.GetTexts(paragraph)));
            }

            return title;
        }

        /// <summary>
        /// Determines whether the given shape is a title.
        /// </summary>
        private static bool IsShapeATitle(Shape shape)
        {
            bool isShape = false;

            var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
            if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type.HasValue)
            {
                switch ((PlaceholderValues)placeholderShape.Type)
                {
                    // Any title shape
                    case PlaceholderValues.Title:
                        isShape = true;
                        break;

                    // A centered title
                    case PlaceholderValues.CenteredTitle:
                        isShape = true;
                        break;
                }
            }

            return isShape;
        }

        /// <summary>
        /// Gets all the notes associated with the slide.
        /// </summary>
        /// <returns>All the notes.</returns>
        /// <see href="http://msdn.microsoft.com/en-us/library/office/gg278319.aspx">Working with Notes Slides</see>
        public IEnumerable<string> GetNotes()
        {
            List<string> notes = new List<string>();
            if (this.slidePart.NotesSlidePart != null)
            {
                foreach (A.Paragraph p in this.slidePart.NotesSlidePart.NotesSlide.Descendants<A.Paragraph>())
                {
                    notes.Add(PptxParagraph.GetTexts(p));
                }
            }
            return notes;
        }

        /// <summary>
        /// Gets all the tables associated with the slide.
        /// </summary>
        /// <returns>All the tables.</returns>
        /// <remarks>Assigns an "artificial" id (tblId) to the tables that match the tag.</remarks>
        public IEnumerable<PptxTable> GetTables()
        {
            List<PptxTable> tables = new List<PptxTable>();

            int tblId = 0;
            foreach (GraphicFrame graphicFrame in this.slidePart.Slide.Descendants<GraphicFrame>())
            {
                var cNvPr = graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties;
                if (cNvPr.Title != null)
                {
                    string title = cNvPr.Title.Value;
                    tables.Add(new PptxTable(this, tblId, title));
                    tblId++;
                }
            }

            return tables;
        }

        /// <summary>
        /// Finds a table given its tag inside the slide.
        /// </summary>
        /// <returns>The table or null.</returns>
        /// <remarks>Assigns an "artificial" id (tblId) to the tables that match the tag.</remarks>
        public IEnumerable<PptxTable> FindTables(string tag)
        {
            List<PptxTable> tables = new List<PptxTable>();
            foreach (PptxTable table in this.GetTables())
            {
                if (table.Title.Contains(tag))
                {
                    tables.Add(table);
                }
            }
            return tables;
        }

        /// <summary>
        /// Type of replacement to perform inside ReplaceTag().
        /// </summary>
        public enum ReplacementType
        {
            /// <summary>
            /// Replaces the tags everywhere.
            /// </summary>
            Global,

            /// <summary>
            /// Does not replace tags that are inside a table.
            /// </summary>
            NoTable
        }

        /// <summary>
        /// Replaces a text (tag) by another inside the slide.
        /// </summary>
        /// <param name="tag">The tag to replace by newText, if null or empty do nothing; tag is a regex string.</param>
        /// <param name="newText">The new text to replace the tag with, if null replaced by empty string.</param>
        /// <param name="replacementType">The type of replacement to perform.</param>
        public void ReplaceTag(string tag, string newText, ReplacementType replacementType)
        {
            foreach (A.Paragraph p in this.slidePart.Slide.Descendants<A.Paragraph>())
            {
                switch (replacementType)
                {
                    case ReplacementType.Global:
                        PptxParagraph.ReplaceTag(p, tag, newText);
                        break;

                    case ReplacementType.NoTable:
                        var tables = p.Ancestors<A.Table>();
                        if (!tables.Any())
                        {
                            // If the paragraph has no table ancestor
                            PptxParagraph.ReplaceTag(p, tag, newText);
                        }
                        break;
                }
            }

            this.Save();
        }

        /// <summary>
        /// Replaces a text (tag) by another inside the slide given a PptxTable.Cell.
        /// This is a convenient method that overloads the original ReplaceTag() method.
        /// </summary>
        /// <param name="tagPair">The tag/new text, BackgroundPicture is ignored.</param>
        /// <param name="replacementType">The type of replacement to perform.</param>
        public void ReplaceTag(PptxTable.Cell tagPair, ReplacementType replacementType)
        {
            this.ReplaceTag(tagPair.Tag, tagPair.NewText, replacementType);
        }

        /// <summary>
        /// Adds a new picture to the slide in order to re-use the picture later on.
        /// </summary>
        internal ImagePart AddPicture(byte[] picture, string contentType)
        {
            ImagePartType type = 0;
            switch (contentType)
            {
                case "image/bmp":
                    type = ImagePartType.Bmp;
                    break;
                case "image/emf": // TODO
                    type = ImagePartType.Emf;
                    break;
                case "image/gif": // TODO
                    type = ImagePartType.Gif;
                    break;
                case "image/ico": // TODO
                    type = ImagePartType.Icon;
                    break;
                case "image/jpeg":
                    type = ImagePartType.Jpeg;
                    break;
                case "image/pcx": // TODO
                    type = ImagePartType.Pcx;
                    break;
                case "image/png":
                    type = ImagePartType.Png;
                    break;
                case "image/tiff": // TODO
                    type = ImagePartType.Tiff;
                    break;
                case "image/wmf": // TODO
                    type = ImagePartType.Wmf;
                    break;
            }

            ImagePart imagePart = this.slidePart.AddImagePart(type);

            // FeedData() closes the stream and we cannot reuse it (ObjectDisposedException)
            // solution: copy the original stream to a MemoryStream
            using (MemoryStream stream = new MemoryStream(picture))
            {
                imagePart.FeedData(stream);
            }

            // No need to detect duplicated images
            // PowerPoint do it for us on the next manual save

            return imagePart;
        }

        /// <summary>
        /// Gets the relationship ID of a given image part.
        /// </summary>
        /// <param name="imagePart">The image part.</param>
        /// <returns>The relationship ID of the image part.</returns>
        internal string GetIdOfImagePart(ImagePart imagePart)
        {
            return this.slidePart.GetIdOfPart(imagePart);
        }

        /// <summary>
        /// Replaces a picture by another inside the slide.
        /// </summary>
        /// <param name="tag">The tag to replace by newPicture, if null or empty do nothing.</param>
        /// <param name="newPicture">The new picture (as a byte array) to replace the tag with, if null do nothing.</param>
        /// <param name="contentType">The picture content type (image/png, image/jpeg...).</param>
        /// <remarks>
        /// <see href="http://stackoverflow.com/questions/7070074/how-can-i-retrieve-images-from-a-pptx-file-using-ms-open-xml-sdk">How can I retrieve images from a .pptx file using MS Open XML SDK?</see>
        /// <see href="http://stackoverflow.com/questions/7137144/how-can-i-retrieve-some-image-data-and-format-using-ms-open-xml-sdk">How can I retrieve some image data and format using MS Open XML SDK?</see>
        /// <see href="http://msdn.microsoft.com/en-us/library/office/bb497430.aspx">How to: Insert a Picture into a Word Processing Document</see>
        /// </remarks>
        public void ReplacePicture(string tag, byte[] newPicture, string contentType)
        {
            if (string.IsNullOrEmpty(tag))
            {
                return;
            }

            if (newPicture == null)
            {
                return;
            }

            ImagePart imagePart = this.AddPicture(newPicture, contentType);

            foreach (Picture pic in this.slidePart.Slide.Descendants<Picture>())
            {
                var cNvPr = pic.NonVisualPictureProperties.NonVisualDrawingProperties;
                if (cNvPr.Title != null)
                {
                    string title = cNvPr.Title.Value;
                    if (title.Contains(tag))
                    {
                        // Gets the relationship ID of the part
                        string rId = this.slidePart.GetIdOfPart(imagePart);

                        pic.BlipFill.Blip.Embed.Value = rId;
                    }
                }
            }

            // Need to save the slide otherwise the relashionship is not saved.
            // Example: <a:blip r:embed="rId2">
            // r:embed is not updated with the right rId
            this.Save();
        }

        /// <summary>
        /// Replaces a picture by another inside the slide.
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="newPictureFile">The new picture (as a file path) to replace the tag with.</param>
        /// <param name="contentType">Type of the content (image/png, image/jpeg...).</param>
        public void ReplacePicture(string tag, string newPictureFile, string contentType)
        {
            byte[] bytes = File.ReadAllBytes(newPictureFile);
            this.ReplacePicture(tag, bytes, contentType);
        }

        /// <summary>
        /// Finds a table (a:tbl) given its "artificial" id (tblId).
        /// </summary>
        /// <param name="tblId">The table id.</param>
        /// <returns>The table or null if not found.</returns>
        /// <remarks>The "artificial" id (tblId) is created inside FindTables().</remarks>
        internal A.Table FindTable(int tblId)
        {
            A.Table tbl = null;

            IEnumerable<GraphicFrame> graphicFrames = this.slidePart.Slide.Descendants<GraphicFrame>();
            GraphicFrame graphicFrame = graphicFrames.ElementAt(tblId);
            if (graphicFrame != null)
            {
                tbl = graphicFrame.Descendants<A.Table>().First();
            }

            return tbl;
        }

        /// <summary>
        /// Removes a table (a:tbl) given its "artificial" id (tblId).
        /// </summary>
        /// <param name="tblId">The table id.</param>
        /// <returns>True if the table has been removed; false otherwise.</returns>
        /// <remarks>
        /// <![CDATA[
        /// p:graphicFrame
        ///  a:graphic
        ///   a:graphicData
        ///    a:tbl (Table)
        /// ]]>
        /// </remarks>
        internal bool RemoveTable(int tblId)
        {
            bool removed = false;

            IEnumerable<GraphicFrame> graphicFrames = this.slidePart.Slide.Descendants<GraphicFrame>();
            GraphicFrame graphicFrame = graphicFrames.ElementAt(tblId);
            if (graphicFrame != null)
            {
                graphicFrame.Remove();
                removed = true;
            }

            return removed;
        }

        /// <summary>
        /// Clones this slide.
        /// </summary>
        /// <returns>The clone.</returns>
        /// <see href="http://blogs.msdn.com/b/brian_jones/archive/2009/08/13/adding-repeating-data-to-powerpoint.aspx">Adding Repeating Data to PowerPoint</see>
        /// <see href="http://startbigthinksmall.wordpress.com/2011/05/17/cloning-a-slide-using-open-xml-sdk-2-0/">Cloning a Slide using Open Xml SDK 2.0</see>
        /// <see href="http://www.exsilio.com/blog/post/2011/03/21/Cloning-Slides-including-Images-and-Charts-in-PowerPoint-presentations-Using-Open-XML-SDK-20-Productivity-Tool.aspx">See Cloning Slides including Images and Charts in PowerPoint presentations and Using Open XML SDK 2.0 Productivity Tool</see>
        public PptxSlide Clone()
        {
            SlidePart template = this.slidePart;

            // Clone slide contents
            SlidePart slidePartClone = this.presentationPart.AddNewPart<SlidePart>();
            using (var templateStream = template.GetStream(FileMode.Open))
            {
                slidePartClone.FeedData(templateStream);
            }

            // Copy layout part
            slidePartClone.AddPart(template.SlideLayoutPart);

            // Copy the image parts
            foreach (ImagePart image in template.ImageParts)
            {
                ImagePart imageClone = slidePartClone.AddImagePart(image.ContentType, template.GetIdOfPart(image));
                using (var imageStream = image.GetStream())
                {
                    imageClone.FeedData(imageStream);
                }
            }

            return new PptxSlide(this.presentationPart, slidePartClone);
        }

        /// <summary>
        /// Inserts this slide after a given target slide.
        /// </summary>
        /// <remarks>This slide will be inserted after the slide specified as a parameter.</remarks>
        /// <see href="http://startbigthinksmall.wordpress.com/2011/05/17/cloning-a-slide-using-open-xml-sdk-2-0/">Cloning a Slide using Open Xml SDK 2.0</see>
        public static void InsertAfter(PptxSlide newSlide, PptxSlide prevSlide)
        {
            // Find the presentationPart
            var presentationPart = prevSlide.presentationPart;

            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the slide id where to insert our slide
            SlideId prevSlideId = null;
            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                // See http://openxmldeveloper.org/discussions/development_tools/f/17/p/5302/158602.aspx
                if (slideId.RelationshipId == presentationPart.GetIdOfPart(prevSlide.slidePart))
                {
                    prevSlideId = slideId;
                    break;
                }
            }

            // Find the highest id
            uint maxSlideId = slideIdList.ChildElements.Cast<SlideId>().Max(x => x.Id.Value);

            // public override T InsertAfter<T>(T newChild, DocumentFormat.OpenXml.OpenXmlElement refChild)
            // Inserts the specified element immediately after the specified reference element.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId + 1;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(newSlide.slidePart);
        }

        /// <summary>
        /// Removes the slide from the PowerPoint file.
        /// </summary>
        /// <see href="http://msdn.microsoft.com/en-us/library/office/cc850840.aspx">How to: Delete a Slide from a Presentation</see>
        public void Remove()
        {
            SlideIdList slideIdList = this.presentationPart.Presentation.SlideIdList;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.RelationshipId == this.presentationPart.GetIdOfPart(this.slidePart))
                {
                    slideIdList.RemoveChild(slideId);
                    break;
                }
            }

            this.presentationPart.DeletePart(this.slidePart);
        }

        /// <summary>
        /// Saves the slide.
        /// </summary>
        /// <remarks>
        /// This is mandatory to save the slides after modifying them otherwise
        /// the next manipulation that will be performed on the pptx won't
        /// include the modifications done before.
        /// </remarks>
        internal void Save()
        {
            this.slidePart.Slide.Save();
        }
    }
}
