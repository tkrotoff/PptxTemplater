using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PptxTemplating.Tests
{
    [TestClass]
    public class PptxTest
    {
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

            pptx.Close();
        }

        [TestMethod]
        public void TestReplaceTagInSlide()
        {
            string srcFileName = "../../files/ReplaceTagInSlide.pptx";
            string dstFileName = "../../files/ReplaceTagInSlide2.pptx";
            File.Copy(srcFileName, dstFileName);

            Pptx pptx = new Pptx(dstFileName, true);
            int nbSlides = pptx.CountSlides();
            Assert.AreEqual(2, nbSlides);

            // First slide
            pptx.ReplaceTagInSlide(0, "{{hello}}", "HELLO HOW ARE YOU?");
            pptx.ReplaceTagInSlide(0, "{{bonjour}}", "BONJOUR TOUT LE MONDE");
            pptx.ReplaceTagInSlide(0, "{{hola}}", "HOLA MAMA QUE TAL?");

            // Second slide
            pptx.ReplaceTagInSlide(1, "{{hello}}", "H");
            pptx.ReplaceTagInSlide(1, "{{bonjour}}", "B");
            pptx.ReplaceTagInSlide(1, "{{hola}}", "H");
            pptx.Close();

            // Check the replaced text is here
            pptx = new Pptx(dstFileName, false);
            nbSlides = pptx.CountSlides();
            StringBuilder result = new StringBuilder();
            for (int i = 0; i < nbSlides; i++)
            {
                string[] texts = pptx.GetAllTextInSlide(i);
                result.Append(string.Join(" ", texts));
                result.Append(" ");
            }
            pptx.Close();
            const string expected = "ReplaceTagInSlide HELLO, world! A tag HOLA inside a sentence A tag BONJOUR inside a sentence HELLO, world! ";
            //Assert.AreEqual(expected, result.ToString());
        }
    }
}
