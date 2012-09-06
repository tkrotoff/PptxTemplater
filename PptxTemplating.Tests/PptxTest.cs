using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PptxTemplating.Tests
{
    [TestClass]
    public class PptxTest
    {
        // See How to: Get All the Text in a Slide in a Presentation http://msdn.microsoft.com/en-us/library/office/cc850836
        // See How to: Get All the Text in All Slides in a Presentation http://msdn.microsoft.com/en-us/library/office/gg278331
        [TestMethod]
        public void TestGetAllTextInAllSlides()
        {
            string file = "../../files/test1.pptx";

            Pptx pptx = new Pptx(file, false);
            int nbSlides = pptx.CountSlides();
            Assert.AreEqual(3, nbSlides);

            var slidesText = new Dictionary<int, string[]>();
            for (int i = 0; i < nbSlides; i++)
            {
                string[] texts = pptx.GetAllTextInSlide(i);
                slidesText.Add(i, texts);
            }

            string[] expected = {"test1", "Hello, world!"};
            CollectionAssert.AreEqual(expected, slidesText[0]);
            expected = new string[]
                           {
                               "Title 1", "Bullet 1", "Bullet 2",
                               "Column 1", "Column 2", "Column 3", "Column 4", "Column 5",
                               "Line 1", "Line 2", "Line 3", "Line 4"
                           };
            CollectionAssert.AreEqual(expected, slidesText[1]);
            expected = new string[] {"Title 2", "Bullet 1", "Bullet 2"};
            CollectionAssert.AreEqual(expected, slidesText[2]);
        }

        [TestMethod]
        public void TestReplaceTagInSlide()
        {
            string srcFileName = "../../files/ReplaceTagInSlide.pptx";
            string dstFileName = "../../files/ReplaceTagInSlide2.pptx";
            File.Copy(srcFileName, dstFileName);

            Pptx pptx = new Pptx(dstFileName, true);
            int nbSlides = pptx.CountSlides();

            for (int i = 0; i < nbSlides; i++)
            {
                pptx.ReplaceTagInSlide(i, "<hello>", "HELLO");
                pptx.ReplaceTagInSlide(i, "<bonjour>", "BONJOUR");
                pptx.ReplaceTagInSlide(i, "<hola>", "HOLA");

                pptx.ReplaceTagInSlide(i, "<notfound>", "NOTFOUND");
            }

            pptx.Close();
        }
    }
}
