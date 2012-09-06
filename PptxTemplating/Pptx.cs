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

        ~Pptx()
        {
            //Close();
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
            IEnumerable<Drawing.Text> texts = slide.Slide.Descendants<Drawing.Text>();
            foreach (Drawing.Text text in texts)
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
            foreach (A.Paragraph paragraph in slide.Slide.Descendants<A.Paragraph>())
            {
                StringBuilder paragraphText = new StringBuilder();

                // Iterate through the lines of the paragraph.
                foreach (A.Text text in paragraph.Descendants<A.Text>())
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

        // Moves a paragraph range in a TextBody shape in the source document
        // to another TextBody shape in the target document.
        // See How to: Move a Paragraph from One Presentation to Another http://msdn.microsoft.com/en-us/library/office/cc850850
        /*public static void MoveParagraphToPresentation(string sourceFile, string targetFile)
        {
            // Open the source file as read/write.
            using (PresentationDocument sourceDoc = PresentationDocument.Open(sourceFile, true))
            {
                // Open the target file as read/write.
                using (PresentationDocument targetDoc = PresentationDocument.Open(targetFile, true))
                {
                    // Get the first slide in the source presentation.
                    SlidePart slide1 = GetFirstSlide(sourceDoc);

                    // Get the first TextBody shape in it.
                    TextBody textBody1 = slide1.Slide.Descendants<TextBody>().First();

                    // Get the first paragraph in the TextBody shape.
                    // Note: "Drawing" is the alias of namespace DocumentFormat.OpenXml.Drawing
                    Drawing.Paragraph p1 = textBody1.Elements<Drawing.Paragraph>().First();

                    // Get the first slide in the target presentation.
                    SlidePart slide2 = GetFirstSlide(targetDoc);

                    // Get the first TextBody shape in it.
                    TextBody textBody2 = slide2.Slide.Descendants<TextBody>().First();

                    // Clone the source paragraph and insert the cloned. paragraph into the target TextBody shape.
                    // Passing "true" creates a deep clone, which creates a copy of the 
                    // Paragraph object and everything directly or indirectly referenced by that object.
                    textBody2.Append(p1.CloneNode(true));

                    // Remove the source paragraph from the source file.
                    textBody1.RemoveChild<Drawing.Paragraph>(p1);

                    // Replace the removed paragraph with a placeholder.
                    textBody1.AppendChild<Drawing.Paragraph>(new Drawing.Paragraph());

                    // Save the slide in the source file.
                    slide1.Slide.Save();

                    // Save the slide in the target file.
                    slide2.Slide.Save();
                }
            }
        }*/

        // Get the slide part of the first slide in the presentation document.
        public static SlidePart GetFirstSlide(PresentationDocument presentationDocument)
        {
            // Get relationship ID of the first slide
            PresentationPart part = presentationDocument.PresentationPart;
            SlideId slideId = part.Presentation.SlideIdList.GetFirstChild<SlideId>();
            string relId = slideId.RelationshipId;

            // Get the slide part by the relationship ID.
            SlidePart slidePart = (SlidePart) part.GetPartById(relId);

            return slidePart;
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
            // Iterate through all the paragraphs in the slide.
            foreach (A.Paragraph p in slide.Slide.Descendants<A.Paragraph>())
            {
                // Iterate through the lines of the paragraph.
                foreach (A.Text t in p.Descendants<A.Text>())
                {
                    if (Regex.Match(t.Text, tag).Success)
                    {
                        string modifiedText = Regex.Replace(t.Text, tag, newText);
                        t.Text = modifiedText;
                    }
                }


                /*if (Regex.Match(p.InnerText, tag).Success)
                {
                    string modifiedText = Regex.Replace(p.InnerText, tag, newText);
                    p.RemoveAllChildren<A.Run>();

                    A.Run r = new A.Run();
                    A.Text t = new A.Text();
                    t.Text = modifiedText;
                    r.Append(t);
                    p.Append(r);

                    //p.AppendChild(new A.Run(new Text(modifiedText)));
                }*/
            }

            // Flush the stream
            //slide.Slide.Save();
        }

        private TextBody GenerateTextBody()
        {
            TextBody textBody1 = new TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties();
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();

            A.Run run1 = new A.Run();
            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US", Dirty = false, SmartTagClean = false };
            A.Text text1 = new A.Text();
            text1.Text = "Hello this is a tag: <hello>";

            run1.Append(runProperties1);
            run1.Append(text1);
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "fr-FR", Dirty = false };

            paragraph1.Append(run1);
            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);
            return textBody1;
        }
    }
}
