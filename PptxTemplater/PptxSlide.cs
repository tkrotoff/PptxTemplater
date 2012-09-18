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
    class PptxSlide
    {
        private readonly SlidePart slidePart;

        public PptxSlide(SlidePart slidePart)
        {
            this.slidePart = slidePart;
        }

        /// <summary>
        /// Gets all text found inside the slide.
        /// </summary>
        /// <remarks>
        /// Some strings inside the array can be empty, this happens when all A.Text from a paragraph are empty
        /// <see href="http://msdn.microsoft.com/en-us/library/office/cc850836">How to: Get All the Text in a Slide in a Presentation</see>
        /// </remarks>
        public string[] GetAllText()
        {
            List<string> texts = new List<string>();
            foreach (A.Paragraph p in this.slidePart.Slide.Descendants<A.Paragraph>())
            {
                texts.Add(PptxParagraph.GetAllText(p));
            }
            return texts.ToArray();
        }

        /// <summary>
        /// Replaces a text (tag) by another inside the slide.
        /// </summary>
        public void ReplaceTag(string tag, string newText)
        {
            /*
             <a:p>
              <a:r>
               <a:rPr lang="en-US" dirty="0" smtClean="0"/>
               <a:t>
                Hello this is a tag: {{hello}}
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
                Another tag: {{bonjour
               </a:t>
              </a:r>
              <a:r>
               <a:rPr lang="en-US" dirty="0" smtClean="0"/>
               <a:t>
                }} le monde !
               </a:t>
              </a:r>
              <a:endParaRPr lang="en-US" dirty="0"/>
             </a:p>
            */

            foreach (A.Paragraph p in this.slidePart.Slide.Descendants<A.Paragraph>())
            {
                PptxParagraph.ReplaceTag(p, tag, newText);
            }
        }

        /// <summary>
        /// Replaces a picture by another inside the slide.
        /// </summary>
        /// <remarks>
        /// <see href="http://stackoverflow.com/questions/7070074/how-can-i-retrieve-images-from-a-pptx-file-using-ms-open-xml-sdk">How can I retrieve images from a .pptx file using MS Open XML SDK?</see>
        /// <see href="http://stackoverflow.com/questions/7137144/how-can-i-retrieve-some-image-data-and-format-using-ms-open-xml-sdk">How can I retrieve some image data and format using MS Open XML SDK?</see>
        /// <see href="http://msdn.microsoft.com/en-us/library/office/bb497430.aspx">How to: Insert a Picture into a Word Processing Document</see>
        /// </remarks>
        public void ReplacePicture(string tag, Stream newPicture, string contentType)
        {
            // FIXME The content type ("image/png", "image/bmp" or "image/jpeg") does not work
            // All files inside the media directory are suffixed with .bin
            // Instead if DocumentFormat.OpenXml.Packaging.ImagePartType is used, files are suffixed with the right extension
            // but I don't want to expose DocumentFormat.OpenXml.Packaging.ImagePartType to the outside world nor
            // want to add boilerplate code if ... else if ... else if ...
            // OpenXML SDK should be fixed and handle "image/png" and friends properly
            ImagePart imagePart = this.slidePart.AddImagePart(contentType);

            // FeedData() closes the stream and we cannot reuse it (ObjectDisposedException)
            // solution: copy the original stream to a MemoryStream
            using (MemoryStream stream = new MemoryStream())
            {
                newPicture.Position = 0;
                newPicture.CopyTo(stream);
                stream.Position = 0;
                imagePart.FeedData(stream);
            }

            // No need to detect duplicated images
            // PowerPoint do it for us on the next manual save

            foreach (Picture pic in this.slidePart.Slide.Descendants<Picture>())
            {
                string xml = pic.NonVisualPictureProperties.OuterXml;

                if (xml.Contains(tag))
                {
                    // Gets the relationship ID of the part
                    string rId = this.slidePart.GetIdOfPart(imagePart);

                    pic.BlipFill.Blip.Embed.Value = rId;
                }
            }
        }

        /// <summary>
        /// Finds a table given its tag inside the slide.
        /// </summary>
        /// <returns>The table or null.</returns>
        public PptxTable[] FindTables(string tag)
        {
            List<PptxTable> tables = new List<PptxTable>();

            foreach (GraphicFrame graphicFrame in this.slidePart.Slide.Descendants<GraphicFrame>())
            {
                string xml = graphicFrame.NonVisualGraphicFrameProperties.OuterXml;

                if (xml.Contains(tag))
                {
                    A.Table tbl = graphicFrame.Descendants<A.Table>().First();
                    tables.Add(new PptxTable(tbl));
                }
            }

            return tables.ToArray();
        }
    }
}
